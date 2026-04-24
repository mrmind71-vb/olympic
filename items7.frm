VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form itemsfrm 
   Caption         =   "»‰Êœ ”œ«œ «·«⁄÷«¡"
   ClientHeight    =   9870
   ClientLeft      =   360
   ClientTop       =   1410
   ClientWidth     =   20370
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   8820
      Width           =   2715
      Begin VB.CheckBox chkAllValue 
         Appearance      =   0  'Flat
         Caption         =   "«·€«¡  ⁄œÌ· ﬂ· «·ﬁÌ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   270
         Width           =   1965
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ÿ»Ì⁄… «·»‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   9270
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1665
      Width           =   2670
      Begin VB.CheckBox xOptional 
         Appearance      =   0  'Flat
         Caption         =   "»‰œ «Œ Ì«—Ì"
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
         Height          =   270
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CheckBox xTax_late_install 
         Appearance      =   0  'Flat
         Caption         =   "›—Êﬁ ﬁÌ„… „÷«›… ··«ﬁ”«ÿ"
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
         Height          =   270
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1260
         Width           =   2325
      End
      Begin VB.CheckBox xTax_late 
         Appearance      =   0  'Flat
         Caption         =   "«·»‰œ ›—Êﬁ ﬁÌ„… „÷«›…"
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
         Height          =   270
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   900
         Width           =   2100
      End
      Begin VB.CheckBox xLate 
         Appearance      =   0  'Flat
         Caption         =   "⁄·ÌÂ €—«„…"
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
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1290
      End
      Begin VB.CheckBox xAllMember 
         Appearance      =   0  'Flat
         Caption         =   "··ﬂ· "
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
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   705
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1185
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   2430
      Visible         =   0   'False
      Width           =   2670
      Begin VB.CheckBox xBasicNew 
         Appearance      =   0  'Flat
         Caption         =   "«”«”Ì ··⁄÷Ê «·ÃœÌœ ›ﬁÿ"
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
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   540
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.CheckBox xBasicDied 
         Appearance      =   0  'Flat
         Caption         =   "√”«”Ì ·√»‰«¡ «·„ Ê›Ì "
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
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   855
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.CheckBox xBasicOld 
         Appearance      =   0  'Flat
         Caption         =   "«”«”Ì ··⁄÷Ê «·ﬁœÌ„ ›ﬁÿ"
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
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "items7.frx":0000
      Height          =   5145
      Left            =   225
      TabIndex        =   12
      Top             =   3645
      Width           =   18870
      _cx             =   33285
      _cy             =   9075
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   8820
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   38
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
         Picture         =   "items7.frx":0013
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items7.frx":21E3
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   39
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
         Picture         =   "items7.frx":432B
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items7.frx":64F3
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   40
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
         Picture         =   "items7.frx":8642
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items7.frx":A822
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   41
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
         Picture         =   "items7.frx":C97D
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "items7.frx":EB39
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   2565
      Top             =   -90
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   10350
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   45
      Width           =   1545
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   45
         Picture         =   "items7.frx":10C88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      Height          =   690
      Left            =   11925
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   45
      Width           =   7215
      Begin VB.CommandButton cmdsave 
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
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "items7.frx":130B2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "items7.frx":15415
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "items7.frx":1798E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "items7.frx":19DFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "items7.frx":1C694
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5970
         Picture         =   "items7.frx":1EC40
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "»‰œ Œ«’ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   45
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CheckBox xShowPaid 
         Appearance      =   0  'Flat
         Caption         =   " ŸÂ— ›Ï «·„ÿ«·»…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   1830
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "€—«„« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6165
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2295
      Width           =   3075
      Begin VB.TextBox xDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
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
         Left            =   495
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   810
         Width           =   675
      End
      Begin VB.CheckBox xMeeting 
         Appearance      =   0  'Flat
         Caption         =   "«·»‰œ €—«„… Ã„⁄Ì… ⁄„Ê„Ì…"
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
         Height          =   240
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   405
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÌÊ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   6
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   855
         Width           =   240
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "€—«„…  √ŒÌ— ”œ«œ »⁄œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   855
         Width           =   1665
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   405
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "‘"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "»‰Êœ «·ﬁ—«»…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   11970
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2115
      Width           =   7125
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   990
         Width           =   375
      End
      Begin VB.CheckBox xIsMember 
         Appearance      =   0  'Flat
         Caption         =   "«·⁄÷Ê ‰›”Â"
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   315
         Width           =   1245
      End
      Begin VB.TextBox xAge2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
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
         Left            =   135
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1035
         Width           =   555
      End
      Begin VB.TextBox xAge1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   555
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   3015
         TabIndex        =   3
         Top             =   630
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
      Begin MSDataListLib.DataCombo xRelation 
         Height          =   330
         Left            =   3015
         TabIndex        =   42
         Top             =   270
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
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   3015
         TabIndex        =   4
         Top             =   990
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
      Begin VB.Label Label1 
         Caption         =   "«·„Ã„Ê⁄…"
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
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1035
         Width           =   990
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
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ ”‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   14
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   675
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ì ”‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   13
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ—Ì» ··⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   11970
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   675
      Width           =   7125
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
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
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   5730
      End
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4635
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
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
         Left            =   4680
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   5
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1035
         Width           =   435
      End
   End
   Begin Threed.SSCommand cmdYear 
      Height          =   375
      Left            =   16965
      TabIndex        =   36
      Top             =   8865
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   661
      _Version        =   196610
      ForeColor       =   192
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«÷€ÿ ·«Œ Ì«— «·Õ«·Ì…"
      ButtonStyle     =   3
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
      Height          =   300
      Left            =   0
      TabIndex        =   49
      Top             =   9570
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.Label xBranchLookup 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "itemsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bedit As Boolean
Dim bEditRecord As Boolean, bAct As Boolean
Public sItem As String
Dim nRound As Long
Dim con As New ADODB.Connection, aTypes As Variant, aSection As Variant
Dim cFilter As String, cFilterLookup As String
Dim oSearch As New Search3, oSearchCode As New Search3, oSearchYear As New Search
Dim formMode As Byte, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub Option1_Click(Index As Integer)
openCardTable
myUndo
End Sub
Private Sub cmdFilter_Click()
'cFilterLookup = ""
'openCardTable
'myUndo
End Sub
Private Sub cmdGroup_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xGroup.BoundText
oFlagfrm.sTable = "FILE6_10G"
oFlagfrm.sCaption = "«·„Ã„Ê⁄…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
DATA3.Recordset.Requery
xGroup.BoundText = sBound
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub

Private Sub cmdSection_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sCode = xGroup.BoundText
oFlagfrm.sTable = "FILE1_10SC"
oFlagfrm.sCaption = "«ﬁ”«„ «·«’‰«›"
oFlagfrm.nZero = -1
oFlagfrm.bedit = bedit
oFlagfrm.Show 1
DATA2.Refresh
If sCode <> "" Then xSection.BoundText = sCode
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub

Private Sub cmdNoBranch_Click()

End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xItem.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xItem.SetFocus
    End If
    Err.Clear
End If
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
Me.Top = 0
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM RELATION_CODES ORDER BY CODE", con)
Set xRelation.RowSource = data1
xRelation.ListField = "Desca"
xRelation.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("SELECT CODE,DESCA FROM GENDER_CODES ORDER BY CODE", con)
Set xGender.RowSource = DATA2
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("SELECT CODE,DESCA FROM FILE6_10G ORDER BY DESCA", con)
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

aTypes = GetRows("select code,desca from PAID_TYPES ORDER BY CODE", con)
aSection = GetRows("select code,desca from TYPE_CODES ORDER BY CODE", con)
Fixgrd

'Set grid1.DataSource = DATA11

'cmdyear.tag = sSeason
cmdYear.Caption = ArbString(retFlag(aSeason, "desca"))
cmdYear.Tag = sSeason

openCardTable
myUndo
End Sub
Private Sub cmdAdd_Click()
mydefine
'xItem.Text = ""
On Error Resume Next
xdesca.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From FILE6_11  Where item = " & MyParn(xItem.text)
    con.Execute "Delete  From FILE6_10  Where item = " & MyParn(xItem.text)
    con.CommitTrans
    If sItem <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "ITEM < " & MyParn(xItem.text), , adSearchBackward, adBookmarkLast
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
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If sItem <> "" Then
    Unload Me
    Exit Sub
End If
If xItem.Tag = DefineMode Then
    cmdAdd_Click
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
xdesca.SetFocus
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
ItemsLookupAll Me, oSearchItem, cFilter
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
bEditRecord = bedit
cmdSave.Enabled = bEditRecord
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode) And bEditRecord
cmdInform.Enabled = (nMode = LoadMode) And Trim(sItem) = ""
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
'cmdFilter.Visible = cFilterLookup <> ""
xItem.Enabled = Not (nMode = LoadMode)
xItem.Tag = nMode
End Sub
Sub mydefine()
xItem.text = Newflag("FILE6_10", "ITEM", con)
xdesca.text = ""
xGroup.BoundText = ""
xValue.text = ""
xTax_late.Value = 0
xTax_late_install.Value = 0
xOptional.Value = 0

xRelation.BoundText = ""
xAge1.text = ""
xAge2.text = ""
xGender.BoundText = "1"

'xNoRate.Value = 0
xLate.Value = 0
xAllMember.Value = 0
xBasicNew.Value = 0
xBasicDied.Value = 0
'xShowPaid.Value = 0
xIsMember.Value = 0
xMeeting.Value = 0
xDays.text = ""
xBasicOld.Value = 0
CellPos 13, 0, grid1.Cols - 1
'CellPos 13, grid1.rows - 2, grid1.Cols - 1
'Fixgrd
Dim Col As Long, Row As Long
For Row = 1 To grid1.rows - 1
    For Col = 2 To grid1.Cols - 1
        grid1.TextMatrix(Row, Col) = IIf(Col < 5, "", 0)
    Next
Next
Handlecontrols DefineMode
StatusBar1.Panels(1).text = "«÷«›… ”Ã· ÃœÌœ"
End Sub
Sub myload()
xItem.text = CardTable!Item
xdesca.text = CardTable!Desca & ""
xValue.text = Myvalue(CardTable!Value)
xGroup.BoundText = CardTable!Group & ""
xTax_late.Value = IIf(CardTable!TAX_LATE, 1, 0)
xTax_late_install.Value = IIf(CardTable!TAX_LATE_INSTALL, 1, 0)
xOptional.Value = IIf(CardTable!optional, 1, 0)

'xShowPaid.Value = IIf(CardTable!showpaid, 1, 0)

xRelation.BoundText = CardTable!relation & ""
xAge1.text = Myvalue(CardTable!Age1)
xAge2.text = Myvalue(CardTable!Age2)
xDays.text = Myvalue(CardTable!days)

xGender.BoundText = CardTable!GENDER & ""
xLate.Value = IIf(CardTable!late, 1, 0)
xAllMember.Value = IIf(CardTable!AllMember, 1, 0)
xBasicNew.Value = IIf(CardTable!BasicNew, 1, 0)
xBasicDied.Value = IIf(CardTable!BasicDied, 1, 0)
xBasicOld.Value = IIf(CardTable!basicOld, 1, 0)
xMeeting.Value = IIf(CardTable!MEETING, 1, 0)
xIsMember.Value = IIf(CardTable!isMember, 1, 0)


myLoadGrd
CalcTotals
CellPos 13, 0, grid1.Cols - 1

grid1_EnterCell

StatusBar1.Panels(1).text = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional Col As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DescA", addstring(xdesca.text))
aInsert = AddFlag(aInsert, "Value", addstring(xValue.text))
aInsert = AddFlag(aInsert, "[GROUP]", addvalue(xGroup.BoundText))
aInsert = AddFlag(aInsert, "Gender", addvalue(xGender.BoundText))
aInsert = AddFlag(aInsert, "relation", addvalue(xRelation.BoundText))
aInsert = AddFlag(aInsert, "Age1", Val(xAge1.text))
aInsert = AddFlag(aInsert, "Age2", Val(xAge2.text))
aInsert = AddFlag(aInsert, "Late", xLate.Value)
aInsert = AddFlag(aInsert, "AllMember", xAllMember.Value)
aInsert = AddFlag(aInsert, "BasicNew", xBasicNew.Value)
aInsert = AddFlag(aInsert, "Tax_late", xTax_late.Value)
aInsert = AddFlag(aInsert, "Tax_late_install", xTax_late_install.Value)
aInsert = AddFlag(aInsert, "[optional]", xOptional.Value)
aInsert = AddFlag(aInsert, "BasicDied", xBasicDied.Value)
aInsert = AddFlag(aInsert, "BasicOld", xBasicOld.Value)
'aInsert = AddFlag(aInsert, "ShowPaid", xShowPaid.Value)
aInsert = AddFlag(aInsert, "Meeting", xMeeting.Value)
aInsert = AddFlag(aInsert, "IsMember", xIsMember.Value)
aInsert = AddFlag(aInsert, "Days", Val(xDays.text))
con.BeginTrans
On Error GoTo myerror
If xItem.Tag = DefineMode Then
    xItem.text = Newflag("FILE6_10", "item", con)
    aInsert = AddFlag(aInsert, "[item]", xItem.text)
    con.Execute addInsert(aInsert, "FILE6_10")
Else
    con.Execute addUpdate(aInsert, "FILE6_10", "item = " & xItem.text)
End If
myreplaceGrd Row, Col
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xItem.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    oSearchItem.Hide
    myUndo
ElseIf ActiveControl.Name = xItem.Name Then
    xItem.text = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
    SendKeys "{TAB}"
    oSearchCode.Hide
    Unload oSearchCode
ElseIf ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 1)
    Grid1_AfterEdit grid1.Row, grid1.Col
    oSearchItem2.Hide
    CellPos 13, grid1.Row, 0
'ElseIf ActiveControl.Name = cmdYear.Name  Then
ElseIf ActiveControl.Name = cmdYear.Name Then
    cmdYear.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    cmdYear.Caption = ArbString(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    Unload oSearchYear
    openCardTable
    myUndo
End If
End Sub
Sub myproc2(pFilter As String)
Unload oSearchItem
cFilterLookup = pFilter
openCardTable
myUndo
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

Private Sub xIgDiscount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CalcTotals
End Sub

Private Sub xItem_LostFocus()
myLostFocus xItem
If xItem.text = "" Then Exit Sub
If Not IsNumeric(xItem.text) Then Exit Sub
CardTable.Find "ITEM = " & xItem.text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    If xItem.Tag = LoadMode Then mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If xdesca.text = "" Then
    If Not bIgMsg Then MsgBox "»Ì«‰ «·’‰› €Ì— „”Ã·"
    Exit Function
End If
'If Trim(xCode.Text) = "" Then
'    If Not bIgMsg Then MsgBox "«·„Ê—œ €Ì— „”Ã·"
'    Exit Function
'End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If IsNumeric(xItem.text) Then
        CardTable.Find "ITEM = " & xItem.text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT FILE6_10.* FROM FILE6_10 "
cFilter = "FILE6_10.OLD = 0"
If cFilterLookup <> "" Then cFilter = cFilter & turn(cFilter, " and ") & cFilterLookup
If sCode <> "" Then cFilter = "FILE6_10.ITEM = " & MyParn(sCode)
If cFilter <> "" Then cString = cString & " WHERE " & cFilter
cString = cString & " ORDER BY FILE6_10.[ITEM]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub CalcTotals()
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And grid1.Row = grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, 1) = "" And grid1.Col = 0 Then
'    KeyAscii = 0
'    If cmdSave.Enabled Then
'        cmdSave_Click
'        CmdAdd_Click
 '   End If
'End If
If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim cWhere As String
If KeyCode = 46 And bEditRecord Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            cWhere = "ITEM = " & xItem.text
            cWhere = cWhere & " AND SECTION = " & grid1.TextMatrix(grid1.Row, 0)
            con.Execute "DELETE FROM FILE6_11 WHERE " & cWhere
            con.CommitTrans
            myLoadGrd
        End If
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Long
'On Error GoTo myerror
With grid1
If Col <> 3 Then
    If chkAllValue.Value = 0 Then
        For i = Row To grid1.rows - 1
            .TextMatrix(i, Col) = .TextMatrix(Row, Col)
        Next
    End If
End If

If Not MYVALID Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If
If myreplace(Row, Col) Then
    xItem.Tag = LoadMode
    xItem.Enabled = False
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
        myLoadGrd
        .ShowCell grid1.rows - 1, 0
    End If
Else
    myLoadGrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
'If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
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
End Sub
Private Sub grid1_EnterCell()
With grid1
If bEditRecord Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub grid1_GotFocus()
'CellPos 13, grid1.Rows - 2, grid1.Cols - 1
grid1_EnterCell
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
ElseIf Col = 2 Then
'    If Not IsNumeric(grid1.EditText) Then
'        Cancel = True
'    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
Dim cFormat As String, i As Long
.Cols = 5
.ColHidden(0) = True
.rows = 1
.TextMatrix(0, 0) = "ﬂÊœ «·‰Ê⁄"
.TextMatrix(0, 1) = "‰Ê⁄ «·⁄÷ÊÌ…"
.TextMatrix(0, 2) = "«·ﬁÌ„…"
.TextMatrix(0, 3) = "«·Œ’„"
.TextMatrix(0, 4) = "ﬁÌ„… „÷«›…"
If IsEmpty(aTypes) Then Exit Sub
.Cols = 5 + UBound(aTypes) + 1
For i = 0 To UBound(aTypes)
    .TextMatrix(0, i + 5) = retFlag(aTypes(i), "desca")
    .ColDataType(i + 5) = flexDTBoolean
Next
.ColWidth(1) = 1600
.RowHeight(0) = 600
For i = 2 To grid1.Cols - 1
    .ColWidth(i) = 892
Next
If IsEmpty(aSection) Then Exit Sub
For i = 0 To UBound(aSection)
    .AddItem ""
    .TextMatrix(i + 1, 0) = retFlag(aSection(i), "code")
    .TextMatrix(i + 1, 1) = retFlag(aSection(i), "desca") & ""
Next
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < 4 Then
    .Col = Col + 1
ElseIf Row < .rows - 1 Then
    .ShowCell Row + 1, 0
    .Select Row + 1, NextEmpty(grid1, Row + 1, 2, 2)
End If
End With
End Sub
Private Sub MyAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Function myreplaceGrd(Row As Long, Col As Long) As Boolean
Dim aInsert As Variant, i As Long, cWhere As String, cInsert As String, cPar As String, cWhereSub As String
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 1, Row)
        cWhere = "file6_11.item = " & xItem.text
        cWhere = cWhere & " and " & "file6_11.section = " & grid1.TextMatrix(i, 0)
        cWhere = cWhere & " and " & "file6_11.year_code = " & cmdYear.Tag
        If EmptyRow(i) Then
            con.Execute "delete from file6_11 where " & cWhere
        Else
            If Col < 5 Then
                aInsert = AddFlag(Empty, "[VALUE]", Val(.TextMatrix(i, 2)))
                aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(.TextMatrix(i, 3)))
                aInsert = AddFlag(aInsert, "[TAX]", Val(.TextMatrix(i, 4)))
                cInsert = cInsert & addUpdate(aInsert, "FILE6_11", cWhere) & ";"
                For i3 = 0 To UBound(aTypes)
                    aInsert = AddFlag(Empty, "[VALUE]", Val(.TextMatrix(i, 2)))
                    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(.TextMatrix(i, 3)))
                    aInsert = AddFlag(aInsert, "[TAX]", Val(.TextMatrix(i, 4)))
                    aInsert = AddFlag(aInsert, "[ITEM]", xItem.text)
                    aInsert = AddFlag(aInsert, "[SECTION]", grid1.TextMatrix(i, 0))
                    aInsert = AddFlag(aInsert, "[TYPE]", retFlag(aTypes(i3), "CODE"))
                    aInsert = AddFlag(aInsert, "[YEAR_CODE]", cmdYear.Tag)
                    cPar = xItem.text & "," & grid1.TextMatrix(i, 0) & "," & cmdYear.Tag & "," & retFlag(aTypes(i3), "CODE")
                    cInsert = cInsert & addInsertTb(aInsert, "FILE6_11", "dbo.doc_type_file6_11(" & cPar & ") = 0") & ";"
                Next
            End If
            If Col = -1 Or Col >= 5 Then
                For i2 = IIf(Col = -1, 5, Col) To IIf(Col = -1, grid1.Cols - 1, Col)
                    cWhereSub = cWhere & " and " & "file6_11.TYPE = " & retFlag(aTypes(i2 - 5), "CODE")
                    aInsert = AddFlag(Empty, "[BASIC]", IIf(Val(.TextMatrix(i, i2)) = 0, "0", "1"))
                    cInsert = cInsert & addUpdate(aInsert, "FILE6_11", cWhereSub) & ";"
                    aInsert = AddFlag(aInsert, "[ITEM]", xItem.text)
                    aInsert = AddFlag(aInsert, "[SECTION]", grid1.TextMatrix(i, 0))
                    aInsert = AddFlag(aInsert, "[YEAR_CODE]", cmdYear.Tag)
                    aInsert = AddFlag(aInsert, "[TYPE]", retFlag(aTypes(i2 - 5), "CODE"))
                    cPar = xItem.text & "," & grid1.TextMatrix(i, 0) & "," & cmdYear.Tag & "," & retFlag(aTypes(i2 - 5), "CODE")
                    cInsert = cInsert & addInsertTb(aInsert, "FILE6_11", "dbo.doc_type_file6_11(" & cPar & ") = 0") & ";"
                Next
            End If
        End If
    Next
    If cInsert <> "" Then con.Execute cInsert
End With
myreplaceGrd = True
End Function
Private Function EmptyRow(Row As Long)
Dim Col As Long
For Col = 2 To grid1.Cols - 1
    If Val(grid1.TextMatrix(Row, Col)) <> 0 Then
        Exit Function
    End If
Next
EmptyRow = True
End Function

Private Sub myLoadGrd()
With grid1
Dim loctable As New ADODB.Recordset, Row As Long, Col As Long, cString As String

For Row = 1 To grid1.rows - 1
    For Col = 2 To .Cols - 1
        .TextMatrix(Row, Col) = IIf(Col < 5, "", 0)
    Next
Next

For i = 0 To UBound(aTypes)
    cField = cField & "," & myiif("TYPE = " & retFlag(aTypes(i), "code"), "CAST(BASIC AS INT)", "0", "MAX") & " AS TYPE" & retFlag(aTypes(i), "CODE")
Next

cString = "SELECT FILE6_11.[SECTION], MAX(value) AS VALUE, MAX(Discount) AS DISCOUNT, MAX(tax) AS tax " & cField & _
         " FROM FILE6_11"
If cmdYear.Tag <> "" Then cString = cString & turn(cString) & "FILE6_11.YEAR_CODE = " & cmdYear.Tag
cString = cString & turn(cString) & "FILE6_11.ITEM = " & xItem.text
cString = cString & " GROUP BY FILE6_11.[SECTION]"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim nFound As Long
Do Until loctable.EOF
    nFound = .FindRow(loctable!Section, , 0)
    If nFound <> -1 Then
        grid1.TextMatrix(nFound, 2) = Myvalue(loctable!Value)
        grid1.TextMatrix(nFound, 3) = Myvalue(loctable!discount)
        grid1.TextMatrix(nFound, 4) = Myvalue(loctable!TAX)
        For i = 4 To loctable.Fields.Count - 1
            grid1.TextMatrix(nFound, i + 1) = IIf(loctable.Fields(i).Value = 0, 0, -1)
        Next
    End If
    loctable.MoveNext
Loop
'cString = "SELECT FILE6_11.CODE,FILE1_10.DESCA,FILE4_10.DESCA,FILE6_11.QUANT,dbo.f_last_cost(FILE6_11.ITEM),FILE6_11.TOTAL,FILE6_11.ID " & _
'          " FROM FILE6_11 INNER JOIN FILE1_10 ON FILE6_11.ITEM = FILE1_10.ITEM LEFT JOIN FILE4_10 ON FILE1_10.CODE = FILE4_10.CODE"
'cString = cString & turn(cString) & "CODE = " & MyParn(xCode.Text) & " ORDER BY FILE6_11.ID"
'DATA11.RecordSource = cString
'DATA11.Refresh
'myAddItem
End With
End Sub
Private Function FoundOtheritem(pGrid As Variant, nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To pGrid.rows - 2
    If i <> nRow Then
        If Trim(pGrid.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub xPrice_LostFocus()
myLostFocus xPrice
CalcTotals
End Sub
Private Sub xRATE_LostFocus()
'Calctotals
'If Val(xCost.Caption) <> 0 And Val(xPrice.Text) <> 0 Then
'    If Round(Val(xRate.Text), nRound) <> Round((Val(xPrice.Text) / Val(xCost.Caption)) - 100, 2) Then
'        xRate.Text = Round(((Val(xPrice.Text) / Val(xCost.Caption)) * 100) - 100, nRound)
'    End If
'Else
'   xRate.Text = ""
'End If
'myLostFocus xRate
End Sub
Private Sub xMotive_GotFocus()
myGotFocus xMotive
End Sub
Private Sub xMotive_LostFocus()
myLostFocus xMotive
CalcTotals
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
CalcTotals
End Sub
Private Sub xCost_sup_GotFocus()
myGotFocus xCost_sup
End Sub
Private Sub xCost_sup_LostFocus()
myLostFocus xCost_sup
If xItem.Tag = DefineMode Then
    If Val(retFlag(aAddress, "FAIR")) <> 0 And Val(xCost_sup.text) <> 0 Then
        xFair.text = Round(Val(xCost_sup.text) * (Val(retFlag(aAddress, "FAIR")) / 100), 2)
    End If
End If
CalcTotals
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xItem_GotFocus()
myGotFocus xItem
End Sub
Private Sub xgroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub cmdYear_Click()
 Set oSearchYear = New Search
Years_LookupAll Me, oSearchYear
End Sub
