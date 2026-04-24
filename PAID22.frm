VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paidfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ì’«·«  ”œ«œ"
   ClientHeight    =   9555
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
   ScaleHeight     =   9555
   ScaleWidth      =   18330
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin Olymbic.CSubclass CSubclass1 
      Left            =   7425
      Top             =   -270
      _ExtentX        =   2011
      _ExtentY        =   1296
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
      Height          =   735
      Index           =   0
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   810
      Width           =   1770
      Begin Threed.SSCommand cmdAddItems 
         Height          =   555
         Left            =   45
         TabIndex        =   56
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
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
         Caption         =   "√÷«›… »‰Êœ «·„ÿ«·»…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5190
      Left            =   180
      TabIndex        =   43
      Top             =   2745
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   9155
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "PAID22.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "PAID22.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "PAID22.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "PAID22.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "grid1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   3
         Left            =   90
         TabIndex        =   44
         Top             =   360
         Width           =   17835
         _cx             =   31459
         _cy             =   8361
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   2
         Left            =   -74910
         TabIndex        =   46
         Top             =   360
         Width           =   17835
         _cx             =   31459
         _cy             =   8361
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   1
         Left            =   -74910
         TabIndex        =   47
         Top             =   360
         Width           =   17835
         _cx             =   31459
         _cy             =   8361
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
         Cols            =   7
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   0
         Left            =   -74910
         TabIndex        =   48
         Top             =   360
         Width           =   17835
         _cx             =   31459
         _cy             =   8361
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
         Cols            =   7
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
      TabIndex        =   38
      Top             =   7920
      Width           =   3615
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   39
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
         Picture         =   "PAID22.frx":0070
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID22.frx":2217
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   40
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
         Picture         =   "PAID22.frx":425E
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID22.frx":6349
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   41
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
         Picture         =   "PAID22.frx":8343
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID22.frx":A454
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   42
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
         Picture         =   "PAID22.frx":C44E
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID22.frx":E672
      End
   End
   Begin VB.CheckBox xCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ".«·”‰… «·Õ«·Ì… ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9900
      RightToLeft     =   -1  'True
      TabIndex        =   14
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
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   -45
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   0
      Width           =   6180
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4995
         TabIndex        =   32
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
         Picture         =   "PAID22.frx":10743
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID22.frx":12B0E
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3735
         TabIndex        =   33
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
         Picture         =   "PAID22.frx":14BB7
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID22.frx":16BBF
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2475
         TabIndex        =   34
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
         Picture         =   "PAID22.frx":18B76
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID22.frx":1B312
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   35
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
         Picture         =   "PAID22.frx":1D7A6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   45
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
         Picture         =   "PAID22.frx":1FAC9
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID22.frx":21E3F
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
      Height          =   2040
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   12795
      Begin VB.TextBox xYears 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6165
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Tag             =   "N"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox xForm_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6165
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10125
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "N"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9765
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
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
         TabIndex        =   3
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   8640
         TabIndex        =   18
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
      Begin VB.Label xdoc_no_zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   180
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label xType_Member 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   765
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ ”‰Ê«  ·„  ”œœ"
         Height          =   240
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1665
         Width           =   1665
      End
      Begin VB.Label xUnPaid_years 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1620
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”‰Ê«  ·„  ”œœ"
         Height          =   240
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1305
         Width           =   1395
      End
      Begin VB.Label xUnPaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   51
         Top             =   1260
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label xType_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1620
         Width           =   5235
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1665
         Width           =   1125
      End
      Begin VB.Label lblYears 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ «·”‰Ê« "
         Height          =   240
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label xLast_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   21
         Top             =   900
         Width           =   3660
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”œ«œ «·⁄÷Ê"
         Height          =   240
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·„ÿ«·»…"
         Height          =   285
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ﬁ”Ì„…"
         Height          =   240
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   930
      End
      Begin VB.Label xYear_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê”„"
         Height          =   240
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1305
         Width           =   1125
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1260
         Width           =   3930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
         TabIndex        =   6
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
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1485
      Width           =   1770
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   36
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "PAID22.frx":23FC2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID22.frx":268E7
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   675
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "PAID22.frx":2913B
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID22.frx":2B29B
      End
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
   Begin VB.Frame FRAME10 
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
      Height          =   1005
      Left            =   12015
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7920
      Width           =   6180
      Begin VB.Label xLate_Total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  „ √Œ—…"
         Height          =   240
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label xTotal_items 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  «·”‰…"
         Height          =   240
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xLate_Items 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ï"
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "€—«„…  √ŒÌ—"
         Height          =   240
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   1050
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   1
      Left            =   1935
      Top             =   1080
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   2
      Left            =   4185
      Top             =   1755
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   3
      Left            =   2475
      Top             =   675
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4410
      Top             =   45
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   0
      Left            =   585
      Top             =   -90
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   63
      Top             =   9090
      Width           =   18330
      _ExtentX        =   32332
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   64
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
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
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   65
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
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
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   66
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
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
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   67
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   68
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
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
      TabIndex        =   61
      Top             =   -45
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label xYear_code3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -405
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   1035
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label xYear_code2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -405
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   675
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label xyear_code1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -405
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   315
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
      TabIndex        =   10
      Top             =   -45
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "paidfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim cList As String
Dim CardTable As ADODB.Recordset, loctable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String, cFilter As String
Dim oSearchDoc As New Search3, oSearchMember As New Search3, oSearchItems As New Search3, oSearchRel As New Search3
Dim bEditRecord As Boolean, bAct As Boolean, aPen As Variant
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, cdef_Box As String
Dim formMode
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function MyReplace(Optional Index As Integer = -1, Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant, i As Integer
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.Text))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(xType.BoundText))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addstring(xYear_code.Caption))
aInsert = AddFlag(aInsert, "[YEAR_CODE1]", addstring(xyear_code1.Caption))
aInsert = AddFlag(aInsert, "[YEAR_CODE2]", addstring(xYear_code2.Caption))
aInsert = AddFlag(aInsert, "[YEAR_CODE3]", addstring(xYear_code3.Caption))
aInsert = AddFlag(aInsert, "[YEARS]", Val(xYears.Text))
aInsert = AddFlag(aInsert, "FORM_NO", addstring(xForm_no.Text))
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERCODE]", "[USERCODE2]"), addvalue(nUsercode))

con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = Newflag("FILE6_20H", "DOC_NO")
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    con.Execute addUpdate(aInsert, "FILE6_20H", "doc_no = " & addstring(xDoc_No.Text))
End If
If Index = -1 Then
    For i = 0 To grid1.UBound
        myreplaceGrd i, -1
    Next
Else
    myreplaceGrd Index, Row
End If
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd(Index As Integer, Row As Long)
Dim aInsert As Variant
With grid1(Index)
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "VALUE", Val(.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "QUANT", Val(.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DISCOUNT_RATE", Val(.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "LATE_RATE", Val(.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(.TextMatrix(i, 11)))
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
    Unload oSearchMember
ElseIf ActiveControl.Name = grid1(0).Name Then
    Dim Index As Integer
    Index = ActiveControl.Index
    If grid1(Index).Col = 0 Then
        grid1(Index).TextMatrix(grid1(Index).Row, 0) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 0)
        grid1(Index).TextMatrix(grid1(idnex).Row, 1) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 1)
        GrdDesc Index, oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 0), grid1(Index).Row
        grid1_AfterEdit Index, grid1(Index).Row, grid1(Index).Col
        Unload oSearchItems
        CellPos Index, 13, grid1(Index).Row, grid1(Index).Col
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
End If
End Sub

Private Sub cmd_closed_Click()
con.BeginTrans
On Error GoTo myerror
con.Execute " update " & cFileHeader & " set CLOSED = " & IIf(xClosed.Value = 1, "0", "1") & " WHERE doc_no = " & MyParn(xDoc_No.Text)
con.CommitTrans
Err.Clear
'openCardTable
myUndo
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 0
oClosefrm.Show 1
'openCardTable
myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1
'openCardTable
myUndo
End Sub
Private Sub cmdAddItems_Click()
myAdditems
End Sub
Private Function myAdditems() As Boolean
Dim nYears As Long, nFirstYear As Integer
If Not ValidInt(xCode.Text) Then
    MsgBox "ﬂÊœ «·⁄÷Ê €Ì— ’ÕÌÕ"
    Exit Function
End If

If xType_Member.Caption = "" Then
    MsgBox "·Ì” ··⁄÷Ê ‰Ê⁄ ⁄÷ÊÌ…"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If

Dim aYear As Variant
aYear = Ret_Year(xDate.Text, , con)
If IsEmpty(aYear) Then
    MsgBox "«· «—ÌŒ €Ì— „‰«”» ·”œ«œ «Ì „Ê”„"
    Exit Function
End If
                
                
If Not xType.MatchedWithList Then
    MsgBox "·« ÌÊÃœ ‰Ê⁄ „ÿ«·»…"
    Exit Function
End If

xYear_code.Caption = ""
xyear_code1.Caption = ""
xYear_code2.Caption = ""
xYear_code3.Caption = ""


Dim aUnPaid As Variant
If findRows(aPaidTypes, "code", xType.BoundText, "is_paid", , False) Then
    aUnPaid = retUnPaid(xCode.Text, retFlag(aYear, "Year"), con, aPaid, aMember)
    
    If retFlag(aUnPaid, "error", False) Then
        MsgBox retFlag(aUnPaid, "desca")
        Exit Function
    End If
            
    If retFlag(aUnPaid, "years") <= 0 Then
        MsgBox "·Ì” ⁄·Ì «·⁄„Ì· ”‰Ê«  ”«»ﬁ…"
        Exit Function
    End If
    'nYears = retFlag(aUnPaid, "Years")
End If
        
Dim aRet As Variant
aRet = addPayment(xCode.Text, myFormat(xDate.Text), xType.BoundText, con)
cInsert = retFlag(aRet, "sql") & ""
If cInsert <> "" Then
    con.Execute cInsert
    xDoc_No.Text = retFlag(aRet, "doc_no")
    'openCardTable
    myUndo
End If
'Dim I As Integer
'For I = 0 To grid1.UBound
'    grid1(I).rows = 1
'    If I > 1 Then
'        SSTab1.TabCaption(I) = ""
'        SSTab1.TabVisible(I) = False
'    End If
'    myAddItem I
'Next
'
'If findRows(aPaidTypes, "code", xType.BoundText, "is_paid", , False) Then
'    xYear_Desca.Caption = retFlag(aYear, "code")
'    nFirstYear = retFlag(aYear, "CODE")
'    For I = 0 To nYears - 1
'        If I = 0 Then
'            xYear_code.Caption = nFirstYear
'        ElseIf I = 1 Then
'            xyear_code1.Caption = nFirstYear - 1
'        ElseIf I = 2 Then
'            xYear_code2.Caption = nFirstYear - 2
'        ElseIf I = 3 Then
'            xYear_code3.Caption = nFirstYear - 3
'        End If
'        SSTab1.TabCaption(I) = Year_Load(nFirstYear - I, "desca")
'        SSTab1.TabVisible(I) = True
'        addPaidItems I, nFirstYear - I
'    Next
'Else
'    xYear_code.Caption = nFirstYear
'    xYear_Desca.Caption = retFlag(aYear, "code")
'    SSTab1.TabCaption(0) = Year_Load(nFirstYear, "desca")
'    SSTab1.TabVisible(0) = True
'    addPaidItems 0, nFirstYear
'End If
End Function
Private Function addPaidItems(Index As Integer, pYear As Integer)
Dim cString As String, nAge As Long, aMember As Variant, bMemberAdd As Boolean
Dim nAll As Long
aMember = Member_Load(xCode.Text, , con)

aPaid = Member_Paid(xCode.Text, , con)
nAll = retAll(aMember)

cString = "SELECT FILE6_11.ITEM,FILE6_10.AGE1,FILE6_10.AGE2 ,FILE6_10.DESCA, FILE6_10.ALLMEMBER, FILE6_10.LATE, FILE6_10.RELATION," & _
      " FILE6_10.ISMEMBER, COALESCE(FILE6_10.AGE1,0), COALESCE(FILE6_10.AGE2,0), FILE6_10.GENDER, " & _
      " FILE6_10.BASICDIED, FILE6_10.BASICNEW,FILE6_10.BASICOLD, FILE6_10.MEETING, FILE6_10.DAYS, FILE6_10.NORATE, " & _
      " FILE6_11.value, FILE6_11.Discount " & _
      " FROM FILE6_10 INNER JOIN FILE6_11 ON FILE6_10.ITEM = FILE6_11.item " & _
      " WHERE FILE6_11.TYPE = " & xType.BoundText & _
      " AND FILE6_11.BASIC = 1 " & _
      " AND FILE6_11.YEAR_CODE = " & pYear & _
      " AND [SECTION] =  " & xType_Member.Caption
cString = cString & " ORDER BY FILE6_10.ITEM"

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With grid1(Index)
If False Then
    .FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "≈Ã„«·Ì|" & "‰”»… Œ’„|" & "ﬁÌ„… Œ’„|" & "≈Ã„«·Ì »⁄œ «·Œ’„|" & "‰”»… €—«„…|" & "ﬁÌ„… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"
End If

bMemberAdd = retFlag(aMember, "Died", False)
Do Until loctable.EOF
    If loctable!isMember Then
        If (Not bMemberAdd) Then
            If AddMemberData(aMember, Index) Then
                .TextMatrix(.rows - 1, 0) = loctable!Item
                .TextMatrix(.rows - 1, 1) = loctable!Desca
                .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
                .TextMatrix(.rows - 1, 3) = "1"
                .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
                If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(Index)
                myAddItem Index
                bMemberAdd = True
            End If
        End If
    ElseIf Not IsNull(loctable!RELATION) Then
        nRelation = addRelation(Index, loctable!RELATION)
        If nRelation > 0 Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = nRelation
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(Index)
            myAddItem Index
            bMemberAdd = True
        End If
    ElseIf loctable!BasicNew Then
        If IsEmpty(aPaid) Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(Index)
        End If
    ElseIf loctable!basicOld Then
        If Not IsEmpty(aPaid) Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(Index)
        End If
    Else
        .TextMatrix(.rows - 1, 0) = loctable!Item
        .TextMatrix(.rows - 1, 1) = loctable!Desca
        .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
        .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
        .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
        If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(Index)
        myAddItem Index
    End If
    loctable.MoveNext
Loop
CalcTotals
End With
End Function
Private Function AddMemberData(aMember As Variant, Index As Variant) As Boolean
Dim nAge As Integer, nGender As Integer
If IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
   nAge = Age(myFormat(retFlag(aMember, "DATE_BIRTH")), myFormat(xDate.Text)) - Index
Else
   nAge = 1
End If

If Val(loctable!Age1 & "") > nAge And Val(loctable!Age1) <> 0 Then Exit Function
If Val(loctable!Age2 & "") < nAge And Val(loctable!Age2 & "") <> 0 Then Exit Function
If (Not IsNull(loctable!GENDER)) Then
    nGender = TurnValue(retFlag(aMember, "Gender", 1), Null, 1)
    If nGender <> loctable!GENDER Then Exit Function
End If
AddMemberData = True
End Function
Private Function addRelation(Index As Integer, nRelation As Integer) As Integer
Dim myRecordSet As New ADODB.Recordset
Dim nAge As Integer, nGender As Integer
cString = " SELECT [CODE],[DATE_BIRTH],COALESCE(GENDER,1) From FILE1_11"
cString = cString & " where relation = " & nRelation
cString = cString & " AND MEMBER = " & xCode.Text
If Not IsNull(loctable!GENDER) Then cString = cString & " AND COALESCE(GENDER,1) = " & loctable!GENDER
myRecordSet.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until myRecordSet.EOF
    If IsDate(myRecordSet!DATE_BIRTH & "") Then
       nAge = Age(myFormat(myRecordSet!DATE_BIRTH), myFormat(xDate.Text)) - Index
    Else
       nAge = 99
    End If
    If (nAge1 >= Val(loctable!Age1 & "") Or Val(loctable!Age1 & "") = 0) And (nAge2 <= Val(loctable!Age2 & "") Or Val(loctable!Age2 & "") = 0) Then
        addRelation = addRelation + 1
    End If
    myRecordSet.MoveNext
Loop
myRecordSet.Close
Set myRecordSet = Nothing
End Function
Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE6_20 where Doc_No = " & xDoc_No.Text
    con.Execute "Delete From FILE6_20H where Doc_No = " & xDoc_No.Text
    con.CommitTrans
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
    
    openCardTable xdoc_no_zero.Caption, "<="
    If CardTable.EOF Then openCardTable , ">"
    If CardTable.EOF Then
        mydefine
    Else
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
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.Text
retAll = retAll + Val(GetField(cString, con) & "")
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(4, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_20H.DOC_NO, FILE6_20H.FORM_NO,PAID_TYPES.DESCA,CONVERT(VARCHAR(10),FILE6_20H.DATE,111),FILE6_20H.YEAR_CODE, FILE1_10.DESCA" & _
          "  FROM  FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_20H.DATE,FILE6_20H.YEAR_CODE,FILE6_20H.Doc_No"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·«” „«—…- «—ÌŒ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE1_10.Desca%% or **FILE6_20H.FORM_NO**" & _
                  " OR ##FILE6_20.Date##)"

listarray(1, 0) = "ﬂÊœ «·⁄÷Ê"
listarray(1, 1) = "(**FILE6_20H.CODE**)"

listarray(2, 0) = "—ﬁ„ «·„” ‰œ"
listarray(2, 1) = "(**FILE6_20H.DOC_NO**)"

listarray(3, 0) = "”‰… «·”œ«œ"
listarray(3, 1) = "(**FILE6_20H.YEAR_CODE**)"

listarray(4, 0) = "‰Ê⁄ «·„ÿ«·»…"
listarray(4, 1) = "(**FILE6_20H.[TYPE]**)"
listarray(4, 2) = "SELECT CODE,DESCA FROM PAID_TYPES"
listarray(4, 3) = "CODE"
listarray(4, 4) = "DESCA"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·«Ì’«·"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "‰Ê⁄ «·„” ‰œ"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(3, 1) = 1350

GrdArray(4, 0) = "”‰… «·”œ«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«·≈”„"
GrdArray(5, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„ «·„ÿ«·»« "
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdNext_Click()
openCardTable xdoc_no_zero.Caption, ">"
If CardTable.EOF Then openCardTable xDoc_No.Text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xdoc_no_zero.Caption, "<"
If CardTable.EOF Then openCardTable xDoc_No.Text, "="
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
'On Error Resume Next
xCode.SetFocus
Err.Clear
End Sub

Private Sub CmdPrint_Click()
doprint
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
'openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub

Private Sub Command1_Click()
Dim loctable As New ADODB.Recordset, aInsert As Variant
loctable.Open "select file1_10.* from file1_10 LEFT JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE WHERE FILE6_20H.CODE IS NULL AND [drop] = 0", con, adOpenStatic, adLockReadOnly, adCmdText
i = Abs(GetField("select min(doc_no) from file6_20h", con))
Do Until loctable.EOF
    i = i + 1
    aInsert = AddFlag(Empty, "[DATE]", addDate("2016-07-01"))
    aInsert = AddFlag(aInsert, "[CODE]", addvalue(loctable!CODE))
    aInsert = AddFlag(aInsert, "[TYPE]", addvalue("1"))
    aInsert = AddFlag(aInsert, "[YEAR_CODE]", addstring("2016"))
    aInsert = AddFlag(aInsert, "[YEARS]", "1")
    aInsert = AddFlag(aInsert, "FORM_NO", -1 * i)
    aInsert = AddFlag(aInsert, "DOC_NO", -1 * i)
    aInsert = AddFlag(aInsert, "OLD", "1")
    con.Execute addInsert(aInsert, "FILE6_20H")
    loctable.MoveNext
Loop
MsgBox "DONE"
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xDoc_No.Tag = LoadMode Then grid1(0).SetFocus Else xCode.SetFocus
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

Set DATA2.Recordset = myRecordSet("select * from paid_types", con)
Set xType.RowSource = DATA2
xType.ListField = "Desca"
xType.BoundColumn = "Code"

aPen = Array(0, 50, 100, 200)

bEdit = True
For i = 0 To 3
    Set grid1(i).DataSource = DATA1(i)
Next
'openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
If Not MYVALID(True) Then
    On Error Resume Next
    grid1(Index).SetFocus
    Err.Clear
    myloadGrd Index
    If Row < grid1(Index).rows - 1 Then
        grid1(Index).Select Row, Col
    Else
        CellPos Index, 13, grid1(Index).rows - 2, grid1(Index).Cols - 1
    End If
    Exit Sub
End If

With grid1(Index)
If Not validRow(Index, Row) Then Exit Sub
If Row = grid1(Index).rows - 1 Then
    myAddItem Index
End If
CalcTotals
If MyReplace(Index, Row) Then
    If xDoc_No.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadGrd
    End If
End If
End With
End Sub
Private Sub grid1_EnterCell(Index As Integer)
With grid1(Index)
    If ((.Col = 0 And .TextMatrix(.Row, .Cols - 1) = "") Or .Col = 2 Or .Col = 3 Or .Col = 5 Or .Col = 8 Or .Col = 11) And bEditRecord Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDKbdMouse
    End If
End With
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not bIgMsg Then
    If grid1(Index).rows < 3 Then
        MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
        Exit Function
    End If
End If

With grid1(Index)
For i = 1 To .rows - 2
'    If .TextMatrix(i, 1) = "" Then
'        .Select i, 0, i, grid1.Cols - 1
'        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
'        Exit Function
'    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
Dim i As Integer
xDoc_No.Text = CardTable!doc_no
xdoc_no_zero.Caption = CardTable!doc_no_zero & ""
xForm_no.Text = CardTable!FORM_NO & ""
xDate.Text = myFormat_p(CardTable!Date)
xCode.Text = CardTable!CODE & ""
xType.BoundText = CardTable!Type & ""
xYears.Text = Myvalue(CardTable!years)
xYear_Desca.Caption = Year_Load(CardTable!year_code, "DESCA_R", con, CardTable!year_code)
xYear_code.Caption = CardTable!year_code
xyear_code1.Caption = CardTable!Year_code1 & ""
xYear_code2.Caption = CardTable!Year_code2 & ""
xYear_code3.Caption = CardTable!Year_code3 & ""

'xSeason.Caption = CardTable!SEASON
xCode_LostFocus
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
Handlecontrols LoadMode
SSTab1.TabCaption(0) = Year_Load(CardTable!year_code & "", "desca", con, CardTable!year_code & "")

myloadGrd 0

If Not IsNull(CardTable!Year_code1) Then
    SSTab1.TabVisible(1) = True
    SSTab1.TabCaption(1) = Year_Load(CardTable!Year_code1, "desca", con, CardTable!Year_code1)
    myloadGrd 1
Else
    SSTab1.TabVisible(1) = False
    SSTab1.TabCaption(1) = ""
    grid1(1).rows = 1
    myAddItem 1
End If

If Not IsNull(CardTable!Year_code2) Then
    SSTab1.TabVisible(2) = True
    SSTab1.TabCaption(2) = Year_Load(CardTable!Year_code2, "desca", con, CardTable!Year_code2)
    myloadGrd 2
Else
    SSTab1.TabVisible(2) = False
    SSTab1.TabCaption(2) = ""
    grid1(2).rows = 1
    myAddItem 2
End If

If Not IsNull(CardTable!Year_code3) Then
    SSTab1.TabVisible(3) = True
    SSTab1.TabCaption(3) = Year_Load(CardTable!Year_code3, "desca", con, CardTable!Year_code3)
    myloadGrd 3
Else
    SSTab1.TabVisible(3) = False
    SSTab1.TabCaption(3) = ""
    grid1(3).rows = 1
    myAddItem 3
End If

CalcTotals


'cmd_closed.BackColor = IIf(CardTable!CLOSED, vbGreen, vbRed)
'cmd_closed.Caption = IIf(CardTable!CLOSED, "„€·ﬁ - › Õ «·„” ‰œ", "„› ÊÕ - ≈€·«ﬁ «·„” ‰œ")
'xusername.Caption = CardTable!UserName & ""
'xusername2.Caption = CardTable!UserName2 & ""
'XTIME1.Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")
'xtime2.Caption = Format(CardTable!Time2, "YYYY-MM-DD HH:NN")

'CellPos index, 13, Grid1.rows - 2, Grid1.Cols - 1
On Error Resume Next
grid1(0).SetFocus
For i = 0 To grid1.UBound
    CellPos i, 13, grid1(i).rows - 2, grid1(Index).Cols - 1
Next
Err.Clear
End Sub
Private Sub myloadGrd(Index As Integer)
Dim cString As String
cString = "SELECT FILE6_20.[ITEM],FILE6_10.DESCA ,FILE6_20.[VALUE],[QUANT],[TOTAL_ITEM],[DISCOUNT_RATE],[DISCOUNT],[TOTAL_DISCOUNT],[LATE_RATE]," & _
          "FILE6_20.[LATE],[TOTAL],FILE6_20.NOTES ,FILE6_20.[ID]" & _
          " From [FILE6_20] INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM WHERE TAB = " & (Index)
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
Set DATA1(Index).Recordset = myRecordSet(cString, con)
Fixgrd Index
myAddItem Index

'cString = "SELECT FILE6_20.CODE,FILE1_20.DESCA,FILE6_20.MEMBER,FILE6_20.MEMBER_SUB,FILE6_20.DESCA,FILE6_20.VALUE,CONVERT(VARCHAR(10),FILE6_20.DATE_BIRTH,111),FILE6_20.NOTES,FILE6_20.[ID] " & _
'           " FROM FILE6_20 INNER JOIN FILE1_20 ON FILE6_20.CODE = FILE1_20.CODE " & _
'           " WHERE FILE6_20.Doc_no = " & MyParn(xDoc_No.Text)
'Data1.Refresh
'myAddItem
'End With
'Calctotals
End Sub
Private Sub mydefine()
Dim i As Integer, aRet As Variant
xDoc_No.Text = Newflag("FILE6_20H", "DOC_NO")
xdoc_no_zero.Caption = ""
'xForm_no.Text = Newflag(cFileHeader, "FORM_NO", con, "SEASON = " & sSeason)
xForm_no.Text = ""
xType.BoundText = "1"
xDate.Text = myFormat_p(Date)
aRet = Ret_Year(xDate.Text, , con)
xYear_Desca.Caption = retFlag(aRet, "desca")
xYear_code.Caption = retFlag(aRet, "code")
xyear_code1.Caption = ""
xYear_code2.Caption = ""
xYear_code3.Caption = ""

xCode.Text = ""
xCodeDesca.Caption = ""
xType_Desca.Caption = ""
xUnPaid.Caption = ""
xUnPaid_years.Caption = ""
xLast_paid.Caption = ""
xType_Member.Caption = ""
'cmd_closed.BackColor = &H8000000F
'cmd_closed.Caption = "-"
'xClosed.Value = 0
'xusername.Caption = ""
'xusername2.Caption = ""
'XTIME1.Caption = ""
'xtime2.Caption = ""

Fixgrd 0
grid1(0).rows = 1
myAddItem 0

SSTab1.TabCaption(0) = Year_Load(xYear_code.Caption, "desca", con, xYear_code.Caption)
For i = 1 To grid1.UBound
    Fixgrd i
    grid1(i).rows = 1
    SSTab1.TabCaption(i) = ""
    SSTab1.TabVisible(i) = False
    myAddItem i
Next

Handlecontrols DefineMode
CalcTotals
On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub
Private Sub LoadTabCaption()
Dim nYear As Integer
If IsDate(xDate.Text) Then
    nYear = Ret_Year(xDate.Text, "code", con, Year(xDate.Text))
    SSTab1.TabCaption(0) = Year_Load(nYear, "desca", con, nYear)
    If SSTab1.TabVisible(1) Then SSTab1.TabCaption(1) = Year_Load(nYear, "desca", con, nYear - 1)
    If SSTab1.TabVisible(2) Then SSTab1.TabCaption(2) = Year_Load(nYear, "desca", con, nYear - 2)
    If SSTab1.TabVisible(3) Then SSTab1.TabCaption(2) = Year_Load(nYear, "desca", con, nYear - 3)
Else
    SSTab1.TabCaption(0) = ""
    SSTab1.TabCaption(1) = ""
    SSTab1.TabCaption(2) = ""
    SSTab1.TabCaption(3) = ""
End If
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit

'cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord

aRecords = retRecords(xdoc_no_zero.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If
cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub grid1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub grid1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then Exit Sub
With grid1(Index)
    If KeyCode = 112 And grid1(Index).Col = 0 Then
        ItemsLookupAll Me, oSearchItems
    ElseIf KeyCode = 13 Then
        CellPos Index, KeyCode, .Row, .Col
    ElseIf KeyCode = 46 And .Row <> .rows - 1 And .rows > 3 And bEditRecord Then
        If MsgBox("Õ–› øø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            con.BeginTrans
            On Error GoTo myerror
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.Execute "Delete from FILE6_20 where ID = " & .TextMatrix(.Row, .Cols - 1)
            End If
            con.CommitTrans
            myRemove Index, .Row
            grid1_EnterCell Index
        End If
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos Index, KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If (grid1(Index).EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    ElseIf Not ValidInt(grid1(Index).EditText) Then
        MsgBox "«·ﬂÊœ €Ì— ”·Ì„"
        Cancel = True
    Else
        If Not GrdDesc(Index, grid1(Index).EditText, Row) Then
           MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ «Ê ·« Ì’·Õ"
           Cancel = True
        End If
    End If
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearchMember
End If
End Sub
Private Sub xCode_LostFocus()
Dim aPaid As Variant, aUnPaid As Variant
myLostFocus xCode
LoadMember
End Sub
Private Sub LoadMember()
xCodeDesca.Caption = ""
xType_Desca.Caption = ""
xLast_paid.Caption = ""
xUnPaid.Caption = ""
xUnPaid_years.Caption = ""

If Not ValidInt(xCode.Text) Then Exit Sub
Dim aMember As Variant
aMember = Member_Load(xCode.Text, , con)
aPaid = Member_Paid(xCode.Text, , con)
If Not IsEmpty(aMember) Then
    xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
    xType_Desca.Caption = retFlag(aMember, "Type_Desca") & ""
    xType_Member.Caption = retFlag(aMember, "type") & ""
End If

If Not IsEmpty(aPaid) Then
    If retFlag(aPaid, "is_save") Then
        xLast_paid.Caption = "Õ«›Ÿ ⁄÷ÊÌ… Õ Ì " & Year_Load(Val(retFlag(aPaid, "year_code") & "") + (Val(retFlag(aPaid, "Years")) - 1), "desca", con, Val(retFlag(aPaid, "year_code") & "") + (Val(retFlag(aPaid, "Years")) - 1))
    Else
        xLast_paid.Caption = "„”œœ Õ Ì " & retFlag(aPaid, "year_desca") & ""
    End If
Else
    xLast_paid.Caption = "·„ Ì”œœ „‰ ﬁ»·"
End If
aUnPaid = retUnPaid(xCode.Text, sSeason, con, aPaid, aMember)
xUnPaid.Caption = retFlag(aUnPaid, "Desca")
xUnPaid_years.Caption = retFlag(aUnPaid, "Years")
End Sub
Private Sub xCurrent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
openCardTable
myUndo
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If Not ValidNum(xDoc_No.Text) Then
     If xDoc_No.Tag = LoadMode Then
        mydefine
    Else
        xDoc_No.Text = Newflag("FILE6_20H", "DOC_NO", con)
    End If
Else
    If (Not (CardTable.EOF)) And xDoc_No.Tag = LoadMode Then
        If CardTable!doc_no = xDoc_No.Text Then
            Exit Sub
        End If
    End If
    
    openCardTable xDoc_No.Text
    If Not CardTable.EOF Then
        myload
    ElseIf xDoc_No.Tag = LoadMode Then
        mydefine
    Else
        xDoc_No.Text = Newflag("FILE6_20H", "DOC_NO", con)
    End If
End If
End Sub
Private Sub xForm_no_LostFocus()
Dim sDoc As String
If Trim(xForm_no.Text) = "" Then Exit Sub
'xDoc_No.Text = GetField("select top 1 doc_no from file6_20h where form_no = " & xForm_no.Text & " and season = " & xSeason.Caption)
'If xDoc_No.Text = "" Then xDoc_No.Text = GetField("select top 1 doc_no from file6_20h where form_no = " & xForm_no.Text)
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
Dim nTotal As Double, Row As Integer, Index As Integer, nTotal_items As Double, nlate_total As Double, nLate_Items As Double
For Index = 0 To 3
    For Row = 1 To grid1(Index).rows - 2
        grid1(Index).TextMatrix(Row, 4) = mRound(Val(grid1(Index).TextMatrix(Row, 2)) * Val(grid1(Index).TextMatrix(Row, 3)), 2)
        grid1(Index).TextMatrix(Row, 6) = mRound(Val(grid1(Index).TextMatrix(Row, 4)) * (Val(grid1(Index).TextMatrix(Row, 5)) / 100), 2)
        grid1(Index).TextMatrix(Row, 7) = Val(grid1(Index).TextMatrix(Row, 4)) - Val(grid1(Index).TextMatrix(Row, 6))
        grid1(Index).TextMatrix(Row, 9) = mRound(Val(grid1(Index).TextMatrix(Row, 8)) * (Val(grid1(Index).TextMatrix(Row, 7)) / 100), 2)
        grid1(Index).TextMatrix(Row, 10) = Val(grid1(Index).TextMatrix(Row, 7)) + Val(grid1(Index).TextMatrix(Row, 9))
        If Index = 0 Then
            nTotal_items = mRound(Val(grid1(Index).TextMatrix(Row, 7)), 2) + nTotal_items
        Else
            nLate_Items = mRound(Val(grid1(Index).TextMatrix(Row, 7)), 2) + nLate_Items
            bOverRide = True
        End If
        nlate_total = mRound(Val(grid1(Index).TextMatrix(Row, 9)), 2) + nlate_total
    Next
Next
xTotal_items.Caption = nTotal_items
xLate_Items.Caption = nLate_Items
xLate_Total.Caption = Myvalue(nlate_total)
xTotal.Caption = nTotal_items + nLate_Items + nlate_total
'StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Myvalue(nTotal, "Fixed")

End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
'If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd(Index As Integer)
With grid1(Index)
.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "≈Ã„«·Ì|" & "‰”»… Œ’„|" & "ﬁÌ„… Œ’„|" & "≈Ã„«·Ì »⁄œ «·Œ’„|" & "‰”»… €—«„…|" & "ﬁÌ„… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"
.ColWidth(0) = 800
.ColWidth(1) = 4000
.ColWidth(2) = 1200
.ColWidth(3) = 1000
.ColWidth(4) = 1300
.ColWidth(5) = 1000
.ColWidth(6) = 800
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(10) = 1300
.ColWidth(11) = 3500
.ColHidden(4) = True
.ColHidden(6) = True
.ColHidden(7) = True
.ColHidden(9) = True

.ColHidden(.Cols - 1) = True
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ColComboList(0) = cList
End With
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String

Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1  FILE6_20H.* FROM FILE6_20H"

If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = "DOC_NO_ZERO " & pSign & addstring(pDoc_No)
End If

'If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_No)

cFilter = ""
cFilter = "FILE6_20H.OLD = 0"
If xCurrent.Value = 1 Then
    Dim aRet As Variant
    aRet = Year_Load(sSeason, , con)
    cFilter = "FILE6_20H.DATE >= " & DateSq(retFlag(aRet, "DATE1"))
    cFilter = cFilter & " AND FILE6_20H.DATE <= " & DateSq(retFlag(aRet, "DATE2"))
End If

'If cmdFilter.Tag <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE6_20H.DOC_NO IN(" & cmdFilter.Tag & ")"
'If chkDay.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "FILE6_20H.[DATE] = " & DateSq(retDate)
'If chkMonth.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & Year(Date) & " AND MONTH(DATE) = " & Month(Date)
'If chkYear.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & Year(Date)
'If cmdYear.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & cmdYear.Tag
'If cmdsup.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "FILE6_20H.CODE  = " & MyParn(oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)) & ")"
If sDoc_no <> "" Then cFilter = "FILE6_20H.DOC_NO = " & addvalue(sDoc_no)

If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by FILE6_20H.doc_no_zero desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by FILE6_20H.doc_no_zero ASC"
End If
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Function retRecords(pDoc_No) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN doc_no_zero <= " & MyParn(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM file6_20H " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If pDoc_No <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xDoc_No.Text) Then
    openCardTable xDoc_No.Text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myAddItem(Index As Integer)
With grid1(Index)
.AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1(Index)
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(Index, OldRow) Then
        myRemove Index, OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Index As Integer, Cancel As Boolean)
If (Not validRow(Index, grid1(Index).Row)) And grid1(Index).Row <> grid1(Index).rows - 1 And grid1(Index).Row <> 0 And grid1(Index).TextMatrix(grid1(Index).Row, grid1(Index).Cols - 1) = "" Then myRemove Index, grid1(Index).Row
End Sub
Private Function validRow(Index As Integer, Row) As Boolean
With grid1(Index)
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(Index As Integer, ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
With grid1(Index)
If Col < .Cols - 2 Then
    .Col = Col + 1 + IIf(Col = 0 Or Col = 3, 1, 0) + IIf(Col = 5, 2, 0) + IIf(Col = 8, 2, 0)
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid1(Index), Row + 1, 0, 3)
    .ShowCell .Row, 0
Else
    .Select Row, Col
End If
End With
End Sub
Private Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For i = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(Row, i)) = "" Then
        NextEmpty = i
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
If IsDate(xDate.Text) Then
    xYear_Desca.Caption = Ret_Year(xDate.Text, "Desca_r", con, Year(xDate.Text))
    LoadTabCaption
Else
'    xYear_Desca.Caption = ""
    LoadTabCaption
End If
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub

Private Sub myRemove(Index As Integer, Row As Long)
grid1(Index).RemoveItem Row
CalcTotals
End Sub
Private Function GrdDesc(Index As Integer, sItem As String, Row As Long) As Boolean
If Trim(sItem) = "" Then Exit Function
Dim aRet As Variant, aMember As Variant
If ValidInt(sItem) Then
    aRet = GetFields("SELECT DESCA,VALUE FROM file6_10 where ITEM = " & sItem)
    grid1(Index).TextMatrix(Row, 1) = retFlag(aRet, "DESCA") & ""
    grid1(Index).TextMatrix(Row, 2) = retFlag(aRet, "VALUE") & ""
    If grid1(Index).TextMatrix(Row, 3) = "" Then
        grid1(Index).TextMatrix(Row, 3) = 1
    End If
End If
GrdDesc = True
End Function
Private Function doprint()
If Not MYVALID Then Exit Function
Dim loctable As New ADODB.Recordset, cString As String
Dim temptable As New ADODB.Recordset
cString = "SELECT FILE6_20.DOC_NO,FILE6_20H.DATE,CASE WHEN USERS.DESCA IS NULL THEN FILE6_20H.USERNAME ELSE USERS.DESCA END AS USER_NAME,FILE6_20H.DATE2,FILE6_20H.CODE AS CODE_MEMBER,FILE1_10.DESCA AS DESCA_MEMBER, FILE6_20.CODE,FILE1_20.DESCA AS ITEM_DESCA,FILE6_20.DESCA," & _
          "FILE6_20.VALUE,FORM_NO,FILE6_20.[NOTES]" & _
          " FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
          " INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE" & _
          " INNER JOIN FILE1_20 ON FILE6_20.CODE = FILE1_20.CODE" & _
          " LEFT JOIN USERS ON FILE6_20H.USERCODE = USERS.CODE"
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & xDoc_No.Text

Dim aTotal As Variant
aTotal = GetFields("Select sum(file6_20.total) as total from file6_20 where doc_no = " & xDoc_No.Text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim i As Long
With loctable
Do Until loctable.EOF
    temptable.AddNew
    i = i + 1
    temptable!str1 = ArbString(Val(loctable!FORM_NO & ""))
    temptable!str2 = ArbString(Format(loctable!Date, "YYYY-MM-DD"))
    temptable!Str3 = TurnValue(ArbString(loctable!CODE_MEMBER))
    temptable!str4 = TurnValue(ArbString(loctable!Desca_Member))
    temptable!str5 = TurnValue(ArbString(Format(loctable!DATE2, "YYYY-MM-DD")))
    temptable!STR6 = Format(Now, "HH:NN")
    temptable!str11 = TurnValue(loctable!Item_Desca)
    temptable!str12 = TurnValue(loctable!Desca)
    temptable!str13 = TurnValue(loctable!notes)
    temptable!str13 = TurnValue(loctable!notes)
    temptable!str14 = TurnValue(loctable!user_name)
    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
    temptable!val1 = Val(loctable!Value & "")
    temptable!Str10 = MyOnly(Val(retFlag(aTotal, "total") & ""))
    temptable!Val10 = i
    temptable.Update
    loctable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
REPORT1.WindowState = crptMaximized
REPORT1.ReportFileName = App.Path & "\Reports\paid.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.ProgressDialog = False
REPORT1.CopiesToPrinter = 1
'REPORT1.Destination = crptToPrinter
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Function

