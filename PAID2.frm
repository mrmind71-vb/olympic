VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paidfrm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ì’«·«  ”œ«œ ‰‘«ÿ —Ì«÷Ï"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18255
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
   ScaleHeight     =   9540
   ScaleWidth      =   18255
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Height          =   600
      Left            =   11385
      Picture         =   "PAID2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   90
      Width           =   1365
   End
   Begin VB.CheckBox xAdded 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   315
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame6 
      Height          =   1725
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   270
      Visible         =   0   'False
      Width           =   1995
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         TabIndex        =   31
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1058
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   990
         TabIndex        =   32
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1614
         _Version        =   196610
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   45
         TabIndex        =   33
         Top             =   135
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
         _Version        =   196610
         ForeColor       =   1118638
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "„” ‰œ „€·ﬁ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   -90
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "PAID2.frx":242A
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PAID2.frx":4BFD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PAID2.frx":71A9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PAID2.frx":9A43
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8550
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   13
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
         Picture         =   "PAID2.frx":BE61
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID2.frx":E031
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
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
         Picture         =   "PAID2.frx":10179
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID2.frx":12341
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
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
         Picture         =   "PAID2.frx":14490
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID2.frx":16670
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
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
         Picture         =   "PAID2.frx":187CB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "PAID2.frx":1A987
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   10710
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   675
      Width           =   7440
      Begin VB.CheckBox xEng 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„Â‰œ”"
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
         Height          =   360
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   495
         Width           =   870
      End
      Begin VB.TextBox xName 
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
         Left            =   135
         MaxLength       =   150
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   6000
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
         Left            =   4320
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   540
         Width           =   1815
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
         Left            =   4320
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ"
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
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   9405
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   1275
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
         Picture         =   "PAID2.frx":1CAD6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "PAID2.frx":1EE39
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   9195
      Width           =   18255
      _ExtentX        =   32200
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
            TextSave        =   "12:24 ’"
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
      Left            =   -405
      Top             =   855
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Left            =   90
      TabIndex        =   4
      Top             =   2025
      Width           =   18060
      _cx             =   31856
      _cy             =   11456
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
      Cols            =   6
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
   Begin VB.Frame Frame9 
      Height          =   645
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   8505
      Visible         =   0   'False
      Width           =   7890
      Begin VB.Label xusercode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   27
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label xUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label XTIME1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label xUserName2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label XTIME2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   2130
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
   Begin VB.Label xBranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   285
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   270
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "paidfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim oSearchDoc As New Search3, oSearchItems As New Search3, oSearchMember As New Search3
Dim bActiviated As Boolean
Dim bEditRecord As Boolean
Dim DocTitle As String
Dim formMode
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Function MyReplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[NAME]", addstring(xName.Text))
aInsert = AddFlag(aInsert, "[ENG]", xEng.Value)
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
aInsert = AddFlag(aInsert, IIf(xDoc_No.Tag = DefineMode, "[USERCODE]", "[USERCODE2]"), addvalue(nUsercode))
con.BeginTrans
'On Error GoTo myError
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Newflag(cFileHeader, "DOC_NO"))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, "doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd Row
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With Grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, Grid1.rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[CODE]", Grid1.TextMatrix(I, 0))
        aInsert = AddFlag(aInsert, "[QUANT]", Grid1.TextMatrix(I, 2))
        aInsert = AddFlag(aInsert, "VALUE", Val(Grid1.TextMatrix(I, 3)))
        If Grid1.TextMatrix(I, Grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_30")
        Else
            con.Execute addUpdate(aInsert, "FILE6_30", "ID = " & Grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.Grid1.TextMatrix(oSearchDoc.Grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xName.Name Then
    xName.Text = oSearchMember.Grid1.TextMatrix(oSearchMember.Grid1.Row, 1)
    Unload oSearchMember
    SendKeys "{tab}"
ElseIf ActiveControl.Name = Grid1.Name Then
    If Grid1.Col = 0 Then
        nFound = FoundOtheritem(Grid1, Grid1.Row, 0, oSearchItems.Grid1.TextMatrix(oSearchItems.Grid1.Row, 0))
        If nFound <> -1 Then
            MsgBox "«·ﬂÊœ „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
            Exit Sub
        End If
    
        Grid1.TextMatrix(Grid1.Row, 0) = oSearchItems.Grid1.TextMatrix(oSearchItems.Grid1.Row, 0)
        Grid1.TextMatrix(Grid1.Row, 1) = oSearchItems.Grid1.TextMatrix(oSearchItems.Grid1.Row, 1)
        GrdDesc Grid1.TextMatrix(Grid1.Row, 0), Grid1.Row
        Grid1_AfterEdit Grid1.Row, Grid1.Col
        Unload oSearchItems
        CellPos 13, Grid1.Row, Grid1.Col
    End If
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
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT  FILE6_30H.DOC_NO,CONVERT(VARCHAR(10),FILE6_30H.DATE,111), FILE6_30H.[NAME],CASE WHEN ENG = 1 THEN '„Â‰œ”' ELSE '⁄÷Ê ‰«œÌ' END" & _
          "  FROM  FILE6_30H  LEFT JOIN USERS ON FILE6_30H.USERCODE = USERS.CODE"

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_30H.DATE,FILE6_30H.Doc_No"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE6_30H.[NAME]%% OR **FILE6_30H.DOC_NO**" & _
                  " OR ##FILE6_30H.Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "‰Ê⁄ «·⁄÷ÊÌ…"
GrdArray(3, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
On Error Resume Next
xName.SetFocus
Err.Clear
End Sub

Private Sub CmdPrint_Click()
doprint
End Sub
Private Sub CmdSave_Click()
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
Private Sub Form_Activate()
On Error Resume Next
If Not bActiviated Then
    bActiviated = True
    If xDoc_No.Tag = LoadMode Then
        Grid1.SetFocus
    Else
        xName.SetFocus
    End If
End If
Err.Clear
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
openCon con
bEdit = True
cFile = "FILE6_30"
cFileHeader = "FILE6_30H"

Set Grid1.DataSource = Data1
Data1.ConnectionString = strCon

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
GRDTABLE.Close
Set CardTable = Nothing
Set GRDTABLE = Nothing
closeCon con
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    Calctotals
    Exit Sub
End If
With Grid1
If Row = Grid1.rows - 1 Then
    myAddItem
End If
Calctotals
If MyReplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.Tag = LoadMode
        xDoc_No.Enabled = False
    End If
    If Grid1.TextMatrix(Row, Grid1.Cols - 1) = "" Then
        myloadgrd
    End If
End If
End With
End Sub
Private Sub Grid1_EnterCell()
If Grid1.Col = 1 Or Grid1.Col = 3 Or Grid1.Col = 4 Or bEditRecord = False Then
    Grid1.Editable = flexEDNone
Else
    Grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_GotFocus()
If Grid1.Row = 0 Then
    Grid1.SetFocus
    Grid1.Select 1, 0
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.Text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Grid1.rows < 3 Then
    MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
    Exit Function
End If

With Grid1
For I = 1 To .rows - 2
    If .TextMatrix(I, 1) = "" Then
        .Select I, 0, I, Grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xt
'xName.Text = CardTable!Name
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
cmd_closed.BackColor = IIf(CardTable!CLOSED, vbGreen, vbRed)
cmd_closed.Caption = IIf(CardTable!CLOSED, "„€·ﬁ - › Õ «·„” ‰œ", "„› ÊÕ - ≈€·«ﬁ «·„” ‰œ")
xusername.Caption = CardTable!UserName & ""
xusername2.Caption = CardTable!UserName2 & ""
XTIME1.Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")
xtime2.Caption = Format(CardTable!Time2, "YYYY-MM-DD HH:NN")
xEng.Value = IIf(CardTable!ENG, 1, 0)
Handlecontrols LoadMode
myloadgrd
CellPos 13, Grid1.rows - 2, Grid1.Cols - 1
On Error Resume Next
Grid1.SetFocus
Err.Clear
End Sub
Private Sub myloadgrd()
With Grid1
cString = "SELECT FILE6_30.CODE,FILE1_30.DESCA,FILE6_30.QUANT,FILE6_30.VALUE,FILE6_30.TOTAL,FILE6_30.[ID] " & _
           " FROM FILE6_30 INNER JOIN FILE1_30 ON FILE6_30.CODE = FILE1_30.CODE" & _
           " WHERE FILE6_30.Doc_no = " & MyParn(xDoc_No.Text)
Data1.RecordSource = cString
Data1.Refresh
myAddItem
End With
Calctotals
Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Newflag(cFileHeader, "DOC_NO"))
xDate.Text = Format(Date, "YYYY-MM-DD")
xName.Text = ""
cmd_closed.BackColor = &H8000000F
cmd_closed.Caption = "-"
xClosed.Value = 0
xusername.Caption = ""
xusername2.Caption = ""
XTIME1.Caption = ""
xtime2.Caption = ""
xEng.Value = 0
Fixgrd
Grid1.rows = 1
myAddItem
Handlecontrols DefineMode
Calctotals
On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit And xClosed.Value = 0
cmd_closed.Enabled = (bEditRecord Or retFlag(aSec, "MANAGER")) And nMode = LoadMode
cmd_CLOSEDDATE.Enabled = retFlag(aSec, "MANAGER")
cmd_open.Enabled = retFlag(aSec, "MANAGER")
cmdNewInv.Enabled = nMode = LoadMode
CmdSave.Enabled = bEditRecord
CmdDelInv.Enabled = nMode = LoadMode And bEditRecord
CmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
CmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
CmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
CmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And doc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then Exit Sub
If KeyCode = 112 And Grid1.Col = 0 Then
    Items_SportLookupAll Me, oSearchItems
ElseIf KeyCode = 13 Then
    CellPos KeyCode, Grid1.Row, Grid1.Col
ElseIf KeyCode = 46 And Grid1.Row <> Grid1.rows - 1 And Grid1.rows > 3 And bEditRecord Then
    If Not MsgBox("Õ–› «·»‰œ", vbOKCancel + vbDefaultButton2) = vbNo Then
        con.BeginTrans
        On Error GoTo myerror
        If Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove Grid1.Row
        Grid1_EnterCell
    End If
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
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If (Grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    ElseIf Not ValidInt(Grid1.EditText) Then
        MsgBox "«·ﬂÊœ €Ì— ”·Ì„"
        Cancel = True
    Else
        nFound = FoundOtheritem(Grid1, Row, 0, Trim(Grid1.EditText))
        If nFound <> -1 Then
            MsgBox "«·ﬂÊœ „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
            Exit Sub
        End If

        If Not GrdDesc(Grid1.EditText, Row) Then
           MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ «Ê ·« Ì’·Õ"
           Cancel = True
        End If
    End If
End If
End Sub
Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    MyLoad
ElseIf xDoc_No.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Function Calctotals()
Dim nTotal As Double
With Grid1
For I = 1 To .rows - 2
    .TextMatrix(I, 4) = Round(Val(.TextMatrix(I, 2)) * Val(.TextMatrix(I, 3)), 2)
    nTotal = nTotal + Round(Val(.TextMatrix(I, 4)), 2)
Next
StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Myvalue(nTotal, "Fixed")
End With
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With Grid1
.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·⁄œœ|" & "«·ﬁÌ„…|" & "«·≈Ã„«·Ì|"
.ColWidth(0) = 800
.ColWidth(1) = 4000
.ColWidth(2) = 1000
.ColWidth(3) = 1200
.ColWidth(4) = 1200
.ColHidden(.Cols - 1) = True
For I = 1 To Grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cFileHeader
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
cString = cString & " Order by " & cFileHeader & ".DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myError
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    MyLoad
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myAddItem()
With Grid1
.AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With Grid1
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
        Calctotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(Grid1.Row)) And Grid1.Row <> Grid1.rows - 1 And Grid1.Row <> 0 And Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) = "" Then myRemove Grid1.Row
End Sub
Private Function validRow(Row) As Boolean
With Grid1
If Trim(xName.Text) = "" Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Not ValidInt(.TextMatrix(Row, 2)) Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < Grid1.Cols - 4 Then
    Grid1.Col = Col + 1 + IIf(Col = 1, 0, 1)
ElseIf Row < Grid1.rows - 1 Then
    Grid1.Select Row + 1, NextEmpty(Grid1, Row + 1, 0, 3)
    Grid1.ShowCell Grid1.Row, 0
Else
    Grid1.Select Row, Col
End If
End Sub
Private Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For I = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(Row, I)) = "" Then
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
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub

Private Sub myRemove(Row As Long)
Grid1.RemoveItem Row
Calctotals
End Sub
Private Function GrdDesc(sItem As String, Row As Long) As Boolean
If Trim(sItem) = "" Then Exit Function
Dim aret As Variant, aMember As Variant
aret = GetFields("SELECT DESCA,VALUE,VALUE2 FROM FILE1_30 where CODE = " & sItem)
Grid1.TextMatrix(Row, 1) = retFlag(aret, "DESCA") & ""
If xEng.Value = 1 Then
    Grid1.TextMatrix(Row, 3) = retFlag(aret, "VALUE") & ""
Else
    Grid1.TextMatrix(Row, 3) = retFlag(aret, "VALUE2") & ""
End If
GrdDesc = True
End Function
Private Function doprint()
If Not MYVALID Then Exit Function
Dim loctable As New ADODB.Recordset, cString As String
Dim temptable As New ADODB.Recordset
cString = "SELECT FILE6_30.DOC_NO,FILE6_30H.DATE,CASE WHEN USERS.DESCA IS NULL THEN FILE6_30H.USERNAME ELSE USERS.DESCA END AS USER_NAME,FILE6_30H.NAME," & _
          "FILE1_30.DESCA AS ITEM_DESCA,FILE6_30.QUANT,FILE6_30.VALUE,FILE6_30.TOTAL" & _
          " FROM FILE6_30 INNER JOIN FILE6_30H ON FILE6_30.DOC_NO = FILE6_30H.DOC_NO " & _
          " INNER JOIN FILE1_30 ON FILE6_30.CODE = FILE1_30.CODE" & _
          " LEFT JOIN USERS ON FILE6_30H.USERCODE = USERS.CODE"
cString = cString & turn(cString) & "FILE6_30.DOC_NO = " & xDoc_No.Text

Dim aTotal As Variant
aTotal = GetFields("Select sum(FILE6_30.total) as total from FILE6_30 where doc_no = " & xDoc_No.Text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim I As Long
With loctable
Do Until loctable.EOF
    temptable.AddNew
    I = I + 1
    temptable!str1 = ArbString(Val(loctable!doc_no & ""))
    temptable!str2 = ArbString(Format(loctable!Date, "YYYY-MM-DD"))
    temptable!str4 = TurnValue(ArbString(loctable!Name))
    temptable!STR6 = Format(Now, "HH:NN")
    temptable!str11 = TurnValue(loctable!Item_Desca)
    temptable!str14 = TurnValue(loctable!user_name)
    temptable!str21 = "≈Ì’«· ”œ«œ ‰‘«ÿ ’Ì›Ì"
    temptable!val1 = Val(loctable!Quant & "")
    temptable!val2 = Val(loctable!Value & "")
    temptable!Val3 = Val(loctable!TOTAL & "")
    
    temptable!Str10 = MyOnly(Val(retFlag(aTotal, "total") & ""))
    
    temptable!Val10 = I
    temptable.Update
    loctable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans

Report1.Reset
Report1.WindowState = crptMaximized
Report1.ReportFileName = App.Path & "\Reports\paid2.rpt"
Report1.DataFiles(0) = tempFile
Report1.ProgressDialog = False
Report1.CopiesToPrinter = 1
'REPORT1.Destination = crptToPrinter
Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Function
Private Sub xEng_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
For I = 1 To Grid1.rows - 2
    GrdDesc Grid1.TextMatrix(I, 0), I
Next
Calctotals
If xDoc_No.Tag = LoadMode Then CmdSave_Click
End Sub

Private Sub xName_GotFocus()
myGotFocus xName
End Sub

Private Sub xName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearchMember
End If
End Sub

Private Sub xName_LostFocus()
myLostFocus xName
End Sub

