VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paidfrm 
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
      Left            =   1845
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   1980
      Width           =   1770
      Begin Threed.SSCommand cmdAddItems 
         Height          =   555
         Left            =   45
         TabIndex        =   57
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
   Begin VB.PictureBox CSubclass1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   64
      Top             =   0
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5190
      Left            =   180
      TabIndex        =   44
      Top             =   2745
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   9155
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "PAID21.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "PAID21.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "PAID21.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "PAID21.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grid1(3)"
      Tab(3).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   3
         Left            =   -74910
         TabIndex        =   45
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
         Index           =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4740
         Index           =   0
         Left            =   90
         TabIndex        =   49
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
      TabIndex        =   39
      Top             =   7965
      Width           =   3570
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2700
         TabIndex        =   40
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
         Picture         =   "PAID21.frx":0070
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID21.frx":2217
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1800
         TabIndex        =   41
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
         Picture         =   "PAID21.frx":425E
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID21.frx":6349
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   945
         TabIndex        =   42
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
         Picture         =   "PAID21.frx":8343
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID21.frx":A454
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   43
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
         Picture         =   "PAID21.frx":C44E
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID21.frx":E672
      End
   End
   Begin VB.CheckBox xCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ".«·”‰… «·Õ«·Ì… ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9855
      RightToLeft     =   -1  'True
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   0
      Width           =   6180
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4995
         TabIndex        =   33
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
         Picture         =   "PAID21.frx":10743
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID21.frx":12B0E
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   3735
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
         Picture         =   "PAID21.frx":14BB7
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID21.frx":16BBF
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2475
         TabIndex        =   35
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
         Picture         =   "PAID21.frx":18B76
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID21.frx":1B312
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   36
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
         Picture         =   "PAID21.frx":1D7A6
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   46
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
         Picture         =   "PAID21.frx":1FAC9
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID21.frx":21E3F
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
      TabIndex        =   6
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
         TabIndex        =   31
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   4
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   8640
         TabIndex        =   19
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
      Begin VB.Label xType_Member 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ ”‰Ê«  ·„  ”œœ"
         Height          =   240
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   1350
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
         RightToLeft     =   -1  'True
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   1620
         Width           =   5235
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1665
         Width           =   1125
      End
      Begin VB.Label lblYears 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ «·”‰Ê« "
         Height          =   240
         Left            =   7155
         RightToLeft     =   -1  'True
         TabIndex        =   32
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
         TabIndex        =   22
         Top             =   900
         Width           =   3660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”œ«œ «·⁄÷Ê"
         Height          =   240
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·„ÿ«·»…"
         Height          =   285
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ﬁ”Ì„…"
         Height          =   240
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê”„"
         Height          =   240
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   1260
         Width           =   3930
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   8
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
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   1485
      Width           =   1725
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   37
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
         Picture         =   "PAID21.frx":23FC2
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID21.frx":268E7
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   38
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
         Picture         =   "PAID21.frx":2913B
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "PAID21.frx":2B29B
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   9015
      Width           =   18330
      _ExtentX        =   32332
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "10:41 „"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   0
      Left            =   2070
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
      TabIndex        =   23
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
         TabIndex        =   58
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
         TabIndex        =   30
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  „ √Œ—…"
         Height          =   240
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  «·”‰…"
         Height          =   240
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ï"
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "€—«„…  √ŒÌ—"
         Height          =   240
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   1050
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   1
      Left            =   4500
      Top             =   765
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
      Left            =   4770
      Top             =   360
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
      Left            =   5940
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
      TabIndex        =   62
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
      TabIndex        =   61
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
      TabIndex        =   60
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
      TabIndex        =   59
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
      TabIndex        =   11
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
Private Function MyReplace(Optional index As Integer = -1, Optional row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant, I As Integer
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
If index = -1 Then
    For I = 0 To grid1.UBound
        myreplaceGrd I, -1
    Next
Else
    myreplaceGrd index, row
End If
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd(index As Integer, row As Long)
Dim aInsert As Variant
'.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "≈Ã„«·Ì|" & "‰”»… Œ’„|" & "ﬁÌ„… Œ’„|" & "≈Ã„«·Ì »⁄œ «·Œ’„|" & "‰”»… €—«„…|" & "ﬁÌ„… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"

With grid1(index)
    For I = IIf(row = -1, 1, row) To IIf(row = -1, .rows - 2, row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "VALUE", Val(.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "QUANT", Val(.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DISCOUNT_RATE", Val(.TextMatrix(I, 5)))
        aInsert = AddFlag(aInsert, "LATE_RATE", Val(.TextMatrix(I, 8)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(.TextMatrix(I, 11)))
        If .TextMatrix(I, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & .TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.row, 0)
    xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.row, 1)
    Unload oSearchMember
ElseIf ActiveControl.Name = grid1(0).Name Then
    Dim index As Integer
    index = ActiveControl.index
    If grid1(index).Col = 0 Then
        grid1(index).TextMatrix(grid1(index).row, 0) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.row, 0)
        grid1(index).TextMatrix(grid1(idnex).row, 1) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.row, 1)
        GrdDesc index, oSearchItems.grid1.TextMatrix(oSearchItems.grid1.row, 0), grid1(index).row
        Grid1_AfterEdit index, grid1(index).row, grid1(index).Col
        Unload oSearchItems
        CellPos index, 13, grid1(index).row, grid1(index).Col
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.row, 0)
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
openCardTable
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
openCardTable
myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1
openCardTable
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
    nYears = retFlag(aUnPaid, "Years")
End If
        
Dim I As Integer
For I = 0 To grid1.UBound
    grid1(I).rows = 1
    If I > 1 Then
        SSTab1.TabCaption(I) = ""
        SSTab1.TabVisible(I) = False
    End If
    myAddItem I
Next

If findRows(aPaidTypes, "code", xType.BoundText, "is_paid", , False) Then
    xYear_Desca.Caption = retFlag(aYear, "code")
    nFirstYear = retFlag(aYear, "CODE")
    For I = 0 To nYears - 1
        If I = 0 Then
            xYear_code.Caption = nFirstYear
        ElseIf I = 1 Then
            xyear_code1.Caption = nFirstYear - 1
        ElseIf I = 2 Then
            xYear_code2.Caption = nFirstYear - 2
        ElseIf I = 3 Then
            xYear_code3.Caption = nFirstYear - 3
        End If
        SSTab1.TabCaption(I) = Year_Load(nFirstYear - I, "desca")
        SSTab1.TabVisible(I) = True
        addPaidItems I, nFirstYear - I
    Next
Else
    xYear_code.Caption = nFirstYear
    xYear_Desca.Caption = retFlag(aYear, "code")
    SSTab1.TabCaption(0) = Year_Load(nFirstYear, "desca")
    SSTab1.TabVisible(0) = True
    addPaidItems 0, nFirstYear
End If
End Function
Private Function addPaidItems(index As Integer, pYear As Integer)
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

With grid1(index)
If False Then
    .FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "≈Ã„«·Ì|" & "‰”»… Œ’„|" & "ﬁÌ„… Œ’„|" & "≈Ã„«·Ì »⁄œ «·Œ’„|" & "‰”»… €—«„…|" & "ﬁÌ„… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"
End If

bMemberAdd = retFlag(aMember, "Died", False)
Do Until loctable.EOF
    If loctable!isMember Then
        If (Not bMemberAdd) Then
            If AddMemberData(aMember, index) Then
                .TextMatrix(.rows - 1, 0) = loctable!Item
                .TextMatrix(.rows - 1, 1) = loctable!Desca
                .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
                .TextMatrix(.rows - 1, 3) = "1"
                .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
                If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(index)
                myAddItem index
                bMemberAdd = True
            End If
        End If
    ElseIf Not IsNull(loctable!Relation) Then
        nRelation = addRelation(index, loctable!Relation)
        If nRelation > 0 Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = nRelation
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(index)
            myAddItem index
            bMemberAdd = True
        End If
    ElseIf loctable!BasicNew Then
        If IsEmpty(aPaid) Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(index)
        End If
    ElseIf loctable!basicOld Then
        If Not IsEmpty(aPaid) Then
            .TextMatrix(.rows - 1, 0) = loctable!Item
            .TextMatrix(.rows - 1, 1) = loctable!Desca
            .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
            .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
            .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
            If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(index)
        End If
    Else
        .TextMatrix(.rows - 1, 0) = loctable!Item
        .TextMatrix(.rows - 1, 1) = loctable!Desca
        .TextMatrix(.rows - 1, 2) = Val(loctable!Value & "")
        .TextMatrix(.rows - 1, 3) = IIf(loctable!AllMember, nAll, 1)
        .TextMatrix(.rows - 1, 5) = Val(loctable!Discount & "")
        If loctable!late Then .TextMatrix(.rows - 1, 8) = aPen(index)
        myAddItem index
    End If
    loctable.MoveNext
Loop
Calctotals
End With
End Function
Private Function AddMemberData(aMember As Variant, index As Variant) As Boolean
Dim nAge As Integer, nGender As Integer
If IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
   nAge = Age(myFormat(retFlag(aMember, "DATE_BIRTH")), myFormat(xDate.Text)) - index
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
Private Function addRelation(index As Integer, nRelation As Integer) As Integer
Dim myRecordSet As New ADODB.Recordset
Dim nAge As Integer, nGender As Integer
cString = " SELECT [CODE],[DATE_BIRTH],COALESCE(GENDER,1) From FILE1_11"
cString = cString & " where relation = " & nRelation
cString = cString & " AND MEMBER = " & xCode.Text
If Not IsNull(loctable!GENDER) Then cString = cString & " AND COALESCE(GENDER,1) = " & loctable!GENDER
myRecordSet.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until myRecordSet.EOF
    If IsDate(myRecordSet!DATE_BIRTH & "") Then
       nAge = Age(myFormat(myRecordSet!DATE_BIRTH), myFormat(xDate.Text)) - index
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
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & xDoc_No.Text, , adSearchBackward, adBookmarkLast
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
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.Text
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
                  " ##FILE6_20.Date##)"

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
Private Sub CmdAdd_Click()
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
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub

Private Sub Command1_Click()

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
Dim I As Long
CSubclass1.SubClassMe SSTab1.hwnd, 0, , vbWhite      '//--- Begin SubClassing
openCon con
Set DATA2.Recordset = myRecordSet("select * from paid_types", con)
Set xType.RowSource = DATA2
xType.ListField = "Desca"
xType.BoundColumn = "Code"

aPen = Array(0, 50, 100, 200)

bEdit = True
For I = 0 To 3
    Set grid1(I).DataSource = DATA1(I)
Next
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub Grid1_AfterEdit(index As Integer, ByVal row As Long, ByVal Col As Long)
If Not MYVALID(True) Then
    On Error Resume Next
    grid1(index).SetFocus
    Err.Clear
    myloadgrd index
    If row < grid1(index).rows - 1 Then
        grid1(index).Select row, Col
    Else
        CellPos index, 13, grid1(index).rows - 2, grid1(index).Cols - 1
    End If
    Exit Sub
End If

With grid1(index)
If Not validRow(index, row) Then Exit Sub
If row = grid1(index).rows - 1 Then
    myAddItem index
End If
Calctotals
If MyReplace(index, row) Then
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.Tag = LoadMode
        xDoc_No.Enabled = False
    End If
    If grid1(index).TextMatrix(row, .Cols - 1) = "" Then
        myloadgrd index
        .ShowCell grid1(index).rows - 1, 0
    End If
End If
End With
End Sub
Private Sub Grid1_EnterCell(index As Integer)
With grid1(index)
    If ((.Col = 0 And .TextMatrix(.row, .Cols - 1) = "") Or .Col = 2 Or .Col = 3 Or .Col = 5 Or .Col = 8 Or .Col = 11) And bEditRecord Then
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

If Not bIsMsg Then
    If grid1(index).rows < 3 Then
        MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
        Exit Function
    End If
End If

With grid1(index)
For I = 1 To .rows - 2
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
Dim I As Integer
xDoc_No.Text = CardTable!doc_no
xForm_no.Text = CardTable!FORM_NO & ""
xDate.Text = myFormat_p(CardTable!Date)
xCode.Text = CardTable!CODE & ""
xType.BoundText = CardTable!Type & ""
xYears.Text = Myvalue(CardTable!years)
xYear_Desca.Caption = Year_Load(CardTable!YEAR_CODE, "DESCA_R", con, CardTable!YEAR_CODE)
xYear_code.Caption = CardTable!YEAR_CODE
xyear_code1.Caption = CardTable!Year_code1 & ""
xYear_code2.Caption = CardTable!Year_code2 & ""
xYear_code3.Caption = CardTable!Year_code3 & ""

'xSeason.Caption = CardTable!SEASON
xCode_LostFocus
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
Handlecontrols LoadMode
SSTab1.TabCaption(0) = Year_Load(CardTable!YEAR_CODE & "", "desca", con, CardTable!YEAR_CODE & "")

myloadgrd 0

If Not IsNull(CardTable!Year_code1) Then
    SSTab1.TabVisible(1) = True
    SSTab1.TabCaption(1) = Year_Load(CardTable!Year_code1, "desca", con, CardTable!Year_code1)
    myloadgrd 1
Else
    SSTab1.TabVisible(1) = False
    SSTab1.TabCaption(1) = ""
    grid1(1).rows = 1
    myAddItem 1
End If

If Not IsNull(CardTable!Year_code2) Then
    SSTab1.TabVisible(2) = True
    SSTab1.TabCaption(2) = Year_Load(CardTable!Year_code2, "desca", con, CardTable!Year_code2)
    myloadgrd 2
Else
    SSTab1.TabVisible(2) = False
    SSTab1.TabCaption(2) = ""
    grid1(2).rows = 1
    myAddItem 2
End If

If Not IsNull(CardTable!Year_code3) Then
    SSTab1.TabVisible(3) = True
    SSTab1.TabCaption(3) = Year_Load(CardTable!Year_code3, "desca", con, CardTable!Year_code3)
    myloadgrd 3
Else
    SSTab1.TabVisible(3) = False
    SSTab1.TabCaption(3) = ""
    grid1(3).rows = 1
    myAddItem 3
End If

Calctotals


'cmd_closed.BackColor = IIf(CardTable!CLOSED, vbGreen, vbRed)
'cmd_closed.Caption = IIf(CardTable!CLOSED, "„€·ﬁ - › Õ «·„” ‰œ", "„› ÊÕ - ≈€·«ﬁ «·„” ‰œ")
'xusername.Caption = CardTable!UserName & ""
'xusername2.Caption = CardTable!UserName2 & ""
'XTIME1.Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")
'xtime2.Caption = Format(CardTable!Time2, "YYYY-MM-DD HH:NN")

'CellPos index, 13, Grid1.rows - 2, Grid1.Cols - 1
On Error Resume Next
grid1(0).SetFocus
For I = 0 To grid1.UBound
    CellPos I, 13, grid1(I).rows - 2, grid1(index).Cols - 1
Next
Err.Clear
End Sub
Private Sub myloadgrd(index As Integer)
Dim cString As String
cString = "SELECT FILE6_20.[ITEM],FILE6_10.DESCA ,FILE6_20.[VALUE],[QUANT],[TOTAL_ITEM],[DISCOUNT_RATE],[DISCOUNT],[TOTAL_DISCOUNT],[LATE_RATE]," & _
          "FILE6_20.[LATE],[TOTAL],FILE6_20.NOTES ,FILE6_20.[ID]" & _
          " From [FILE6_20] INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM WHERE TAB = " & (index)
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
Set DATA1(index).Recordset = myRecordSet(cString, con)
Fixgrd index
myAddItem index

'cString = "SELECT FILE6_20.CODE,FILE1_20.DESCA,FILE6_20.MEMBER,FILE6_20.MEMBER_SUB,FILE6_20.DESCA,FILE6_20.VALUE,CONVERT(VARCHAR(10),FILE6_20.DATE_BIRTH,111),FILE6_20.NOTES,FILE6_20.[ID] " & _
'           " FROM FILE6_20 INNER JOIN FILE1_20 ON FILE6_20.CODE = FILE1_20.CODE " & _
'           " WHERE FILE6_20.Doc_no = " & MyParn(xDoc_No.Text)
'Data1.Refresh
'myAddItem
'End With
'Calctotals
End Sub
Private Sub mydefine()
Dim I As Integer, aret As Variant
xDoc_No.Text = Newflag("FILE6_20H", "DOC_NO")
'xForm_no.Text = Newflag(cFileHeader, "FORM_NO", con, "SEASON = " & sSeason)
xForm_no.Text = ""
xType.BoundText = "1"
xDate.Text = myFormat_p(Date)
aret = Ret_Year(xDate.Text, , con)
xYear_Desca.Caption = retFlag(aret, "desca")
xYear_code.Caption = retFlag(aret, "code")
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
For I = 1 To grid1.UBound
    Fixgrd I
    grid1(I).rows = 1
    SSTab1.TabCaption(I) = ""
    SSTab1.TabVisible(I) = False
    myAddItem I
Next

Handlecontrols DefineMode
Calctotals
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
bEditRecord = bEdit And xClosed.Value = 0
xCode.Enabled = bEditRecord And nMode = DefineMode
'cmd_closed.Enabled = (bEditRecord Or retFlag(aSec, "MANAGER")) And nMode = LoadMode
'cmd_CLOSEDDATE.Enabled = retFlag(aSec, "MANAGER")
'cmd_open.Enabled = retFlag(aSec, "MANAGER")
xYears.Enabled = Val(xType.BoundText) = 2
cmdAdd.Enabled = nMode = LoadMode
cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And doc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub grid1_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub grid1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then Exit Sub
With grid1(index)
    If KeyCode = 112 And grid1(index).Col = 0 Then
        ItemsLookupAll Me, oSearchItems
    ElseIf KeyCode = 13 Then
        CellPos index, KeyCode, .row, .Col
    ElseIf KeyCode = 46 And .row <> .rows - 1 And .rows > 3 And bEditRecord Then
        If MsgBox("Õ–› øø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            con.BeginTrans
            On Error GoTo myerror
            If .TextMatrix(.row, .Cols - 1) <> "" Then
                con.Execute "Delete from FILE6_20 where ID = " & .TextMatrix(.row, .Cols - 1)
            End If
            con.CommitTrans
            myRemove index, .row
            Grid1_EnterCell index
        End If
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(index As Integer, ByVal row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos index, KeyCode, row, Col
End Sub
Private Sub grid1_ValidateEdit(index As Integer, ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If (grid1(index).EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    ElseIf Not ValidInt(grid1(index).EditText) Then
        MsgBox "«·ﬂÊœ €Ì— ”·Ì„"
        Cancel = True
    Else
        If Not GrdDesc(index, grid1(index).EditText, row) Then
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
If Trim(xDoc_No.Text) = "" Then
    If xDoc_No.Tag = LoadMode Then mydefine
    Exit Sub
End If
CardTable.Find "Doc_no = " & xDoc_No.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xDoc_No.Tag = LoadMode Then
    mydefine
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
Private Function Calctotals(Optional bOverRide As Boolean = False)
Dim nTotal As Double, row As Integer, index As Integer, nTotal_items As Double, nlate_total As Double, nLate_Items As Double
For index = 0 To 3
    For row = 1 To grid1(index).rows - 2
        grid1(index).TextMatrix(row, 4) = mRound(Val(grid1(index).TextMatrix(row, 2)) * Val(grid1(index).TextMatrix(row, 3)), 2)
        grid1(index).TextMatrix(row, 6) = mRound(Val(grid1(index).TextMatrix(row, 4)) * (Val(grid1(index).TextMatrix(row, 5)) / 100), 2)
        grid1(index).TextMatrix(row, 7) = Val(grid1(index).TextMatrix(row, 4)) - Val(grid1(index).TextMatrix(row, 6))
        grid1(index).TextMatrix(row, 9) = mRound(Val(grid1(index).TextMatrix(row, 8)) * (Val(grid1(index).TextMatrix(row, 7)) / 100), 2)
        grid1(index).TextMatrix(row, 10) = Val(grid1(index).TextMatrix(row, 7)) + Val(grid1(index).TextMatrix(row, 9))
        If index = 0 Then
            nTotal_items = mRound(Val(grid1(index).TextMatrix(row, 7)), 2) + nTotal_items
        Else
            nLate_Items = mRound(Val(grid1(index).TextMatrix(row, 7)), 2) + nLate_Items
            bOverRide = True
        End If
        nlate_total = mRound(Val(grid1(index).TextMatrix(row, 9)), 2) + nlate_total
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
Private Sub Fixgrd(index As Integer)
With grid1(index)
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
For I = 1 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColComboList(0) = cList
End With
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE6_20H.* FROM FILE6_20H"
cFilter = ""
If xCurrent.Value = 1 Then
    Dim aret As Variant
    aret = Year_Load(sSeason, , con)
    cFilter = "FILE6_20H.DATE >= " & DateSq(retFlag(aret, "DATE1"))
    cFilter = cFilter & " AND FILE6_20H.DATE >= " & DateSq(retFlag(aret, "DATE2"))
    'cFilter = "YEAR_CODE + (YEARS - 1) = " & sSeason
End If
If sDoc_no <> "" Then cFilter = " DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " Order by FILE6_20H.DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
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
Private Sub myAddItem(index As Integer)
With grid1(index)
.AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1(index)
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(index, OldRow) Then
        myRemove index, OldRow
        Calctotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(index As Integer, Cancel As Boolean)
If (Not validRow(index, grid1(index).row)) And grid1(index).row <> grid1(index).rows - 1 And grid1(index).row <> 0 And grid1(index).TextMatrix(grid1(index).row, grid1(index).Cols - 1) = "" Then myRemove index, grid1(index).row
End Sub
Private Function validRow(index As Integer, row) As Boolean
With grid1(index)
If Trim(.TextMatrix(row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(index As Integer, ByRef KeyCode, ByVal row As Long, ByVal Col As Long)
KeyCode = 0
With grid1(index)
If Col < .Cols - 2 Then
    .Col = Col + 1 + IIf(Col = 0 Or Col = 3, 1, 0) + IIf(Col = 5, 2, 0) + IIf(Col = 8, 2, 0)
ElseIf row < .rows - 1 Then
    .Select row + 1, NextEmpty(grid1(index), row + 1, 0, 3)
    .ShowCell .row, 0
Else
    .Select row, Col
End If
End With
End Sub
Private Function NextEmpty(pGrid As Object, row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For I = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(row, I)) = "" Then
        NextEmpty = I
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

Private Sub myRemove(index As Integer, row As Long)
grid1(index).RemoveItem row
Calctotals
End Sub
Private Function GrdDesc(index As Integer, sItem As String, row As Long) As Boolean
If Trim(sItem) = "" Then Exit Function
Dim aret As Variant, aMember As Variant
If ValidInt(sItem) Then
    aret = GetFields("SELECT DESCA,VALUE FROM file6_10 where ITEM = " & sItem)
    grid1(index).TextMatrix(row, 1) = retFlag(aret, "DESCA") & ""
    grid1(index).TextMatrix(row, 2) = retFlag(aret, "VALUE") & ""
    If grid1(index).TextMatrix(row, 3) = "" Then
        grid1(index).TextMatrix(row, 3) = 1
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

Dim I As Long
With loctable
Do Until loctable.EOF
    temptable.AddNew
    I = I + 1
    temptable!str1 = ArbString(Val(loctable!FORM_NO & ""))
    temptable!str2 = ArbString(Format(loctable!Date, "YYYY-MM-DD"))
    temptable!Str3 = TurnValue(ArbString(loctable!CODE_MEMBER))
    temptable!str4 = TurnValue(ArbString(loctable!Desca_Member))
    temptable!str5 = TurnValue(ArbString(Format(loctable!date2, "YYYY-MM-DD")))
    temptable!STR6 = Format(Now, "HH:NN")
    temptable!str11 = TurnValue(loctable!Item_Desca)
    temptable!str12 = TurnValue(loctable!Desca)
    temptable!str13 = TurnValue(loctable!notes)
    temptable!str13 = TurnValue(loctable!notes)
    temptable!str14 = TurnValue(loctable!user_name)
    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
    temptable!val1 = Val(loctable!Value & "")
    temptable!Str10 = MyOnly(Val(retFlag(aTotal, "total") & ""))
    temptable!Val10 = I
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

