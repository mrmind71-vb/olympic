VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form boxMovefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16935
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   16935
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "ÿ»«⁄… «Ã„«·Ì Õ—ﬂ… ÌÊ„Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8370
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9630
      Width           =   2400
   End
   Begin VB.CommandButton cmdPrint1 
      Caption         =   "ÿ»«⁄… «Ã„«·Ì Õ—ﬂ… «·Œ“Ì‰…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5895
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   9630
      Width           =   2445
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   6750
      Top             =   -45
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
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
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   225
      Width           =   4650
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2250
         Picture         =   "boxmove.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "boxmove.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3375
         Picture         =   "boxmove.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1170
         Picture         =   "boxmove.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   10575
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6270
      Begin VB.TextBox XDATE2 
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
         Left            =   2025
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
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
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XCODE 
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
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·Œ“‰… :"
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
         Height          =   270
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   945
      End
      Begin VB.Label xDesca 
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
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   3300
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8565
      Left            =   5895
      TabIndex        =   8
      Top             =   990
      Width           =   10995
      _cx             =   19394
      _cy             =   15108
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
      Cols            =   8
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   3795
      Left            =   135
      TabIndex        =   15
      Top             =   45
      Width           =   5685
      _cx             =   10028
      _cy             =   6694
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
      Rows            =   50
      Cols            =   5
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame5 
      Height          =   6450
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3825
      Width           =   5685
      Begin VB.Label xPur_cash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   585
         Width           =   1320
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‘ —Ì«  ‰ﬁœÌ…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label xWages 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   4185
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„— »« "
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   4275
         Width           =   510
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”·›"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   4590
         Width           =   360
      End
      Begin VB.Label xInstall_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   4545
         Width           =   1320
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—œ ”·›"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   4590
         Width           =   585
      End
      Begin VB.Label xinstall_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   4545
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " „œ›Ê⁄«  „œÌ‰Ê‰"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   3510
         Width           =   1305
      End
      Begin VB.Label xCredit_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   3465
         Width           =   1320
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ „œÌ‰Ê‰"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   3510
         Width           =   915
      End
      Begin VB.Label xCredit_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   3465
         Width           =   1275
      End
      Begin VB.Label xPart_In 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   3825
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ìœ«⁄ ‘—ﬂ«¡"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   3870
         Width           =   900
      End
      Begin VB.Label xDebit_In 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   3105
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ·œ«∆‰Ê‰"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   3150
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ê—«ﬁ œ›⁄"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label xChq_Out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1665
         Width           =   1320
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁ»Ê÷«  œ«∆‰Ê‰"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   3150
         Width           =   1260
      End
      Begin VB.Label xDebit_Out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   3105
         Width           =   1320
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÊÌ«  «·Ì"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2790
         Width           =   885
      End
      Begin VB.Label xTrust_cash_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2745
         Width           =   1275
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "⁄Âœ ··Œ“‰…"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2430
         Width           =   840
      End
      Begin VB.Label xTrust_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   2385
         Width           =   1275
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «·Œ“‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   5310
         Width           =   945
      End
      Begin VB.Label xLast_Balance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   5265
         Width           =   1275
      End
      Begin VB.Label xTotal_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   4905
         Width           =   1320
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì ’«œ—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   4950
         Width           =   990
      End
      Begin VB.Label xTrust_cash_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2745
         Width           =   1320
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÊÌ«  „‰"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2790
         Width           =   840
      End
      Begin VB.Label xTrust_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2385
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "⁄Âœ „‰ «·Œ“‰…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2430
         Width           =   1110
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”ÕÊ»«  ‘—ﬂ«¡"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   3915
         Width           =   1245
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕÊÌ· „‰ «·Œ“Ì‰…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1350
         Width           =   1320
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ìœ«⁄«  »‰ﬂÌ…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„’«—Ì› ⁄«„…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2070
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄«  „Ê—œÌ‰"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label xTrans_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1305
         Width           =   1320
      End
      Begin VB.Label xBank_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label xCash_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label xPart_Out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3825
         Width           =   1320
      End
      Begin VB.Label xCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2025
         Width           =   1320
      End
      Begin VB.Label xTotal_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   4905
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì Ê«—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   4950
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2745
         Width           =   60
      End
      Begin VB.Label xIncome 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2025
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì—«œ«  «Œ—Ì"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2070
         Width           =   1035
      End
      Begin VB.Label xChq_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1665
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ê—«ﬁ ﬁ»÷"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1710
         Width           =   885
      End
      Begin VB.Label xTrans_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1305
         Width           =   1275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕÊÌ· ··Œ“‰Ì…"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1350
         Width           =   1050
      End
      Begin VB.Label xBank_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”ÕÊ»«  »‰ﬂÌ…"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label xCash_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁ»Ê÷«  ⁄„·«¡"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label xFirst_Balance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ ”«»ﬁ"
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   855
      End
   End
End
Attribute VB_Name = "boxMovefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, oSearchBox As New Search3
Public Sub myloadgrd()
Dim cString As String, loctable As New ADODB.Recordset, nBalance As Double, nPrevious As Double
With grid1
grid1.Rows = 1
If IsDate(xdate1.Text) Then
   cString = "Select sum([PLUS] - MINUS) from BOXMOVE WHERE BOX = " & MyParn(XCODE.Text) & _
              " AND DATE < " & DateSq(xdate1.Text)
   nPrevious = Round(Val(GetField(cString) & ""), 2)
   If nPrevious <> 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
        .TextMatrix(.Rows - 1, 4) = Round(nPrevious, 2)
        .TextMatrix(.Rows - 1, 6) = Round(nPrevious, 2)
   End If
End If

cString = "select BOXMOVE.*  " & _
          " From BOXMOVE"
cString = cString & turn(cString) & "BOX = " & MyParn(XCODE.Text)

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & "date >= " & DateSq(xdate1.Text)
End If

If IsDate(XDATE2.Text) Then
    cString = cString & turn(cString) & "Date <= " & DateSq(XDATE2.Text)
End If
cString = cString & " Order by BOXMOVE.date,BOXMOVE.FLAG,BOXMOVE.DOC_NO,BOXMOVE.PLUS DESC"
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext
    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Round(Val(loctable!PLUS & ""), 2) - Round(Val(loctable!MINUS & ""), 2)
        .TextMatrix(.Rows - 1, 0) = loctable!desca & ""
        .TextMatrix(.Rows - 1, 1) = loctable!CodeDesca & ""
        .TextMatrix(.Rows - 1, 2) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 3) = loctable!DOC_NO & ""
        .TextMatrix(.Rows - 1, 4) = Myvalue(Round(Val(loctable!PLUS & ""), 2))
        .TextMatrix(.Rows - 1, 5) = Myvalue(Round(Val(loctable!MINUS & ""), 2))
        .TextMatrix(.Rows - 1, 6) = Round(nPrevious, 2)
        .TextMatrix(.Rows - 1, 7) = loctable!flag
        loctable.MoveNext
    Loop
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 4, "#", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 5, "#", vbYellow, vbRed, True, "  "
    If grid1.Rows > 1 Then
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 1) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 2) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 3) = "«·«Ã„«·Ì"
        .MergeCells = flexMergeRestrictRows
        .MergeRow(grid1.Rows - 1) = True
        .TextMatrix(.Rows - 1, 6) = Round(nPrevious, 2)
    End If
End With
End Sub
Private Sub MyLoadTotal()
Dim sourcetable As New ADODB.Recordset, nBalance As Double

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
End If

If XCODE.Text <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XCODE.Text)
End If

'--------------  Ê«—œ
If IsDate(xdate1.Text) Then
    cField1 = "(" & _
               "Select Sum(PLUS - MINUS) From BoxMove " & _
               turn(cwhere) & cwhere & _
               ") as First_Balance"
Else
    cField1 = "0 " & _
              " as First_Balance"
End If

cField1 = cField1 & "," & _
        myiif( _
        "(FLAG = 0)", "PLUS - MINUS") & _
        " As First_Bal"
                                  
cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 1 or Flag = 4 )", "PLUS - MINUS") & _
        " As Cash_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 14)", "PLUS") & _
        " As Bank_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 8)", "PLUS") & _
        " As Trans_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 6)", "PLUS ") & _
        " As Income"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 13)", "PLUS") & _
        " As Chq_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 103)", "PLUS") & _
        " As Trust_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 17)", "PLUS") & _
        " As Part_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 19)", "PLUS") & _
        " As Debit_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 21)", "PLUS") & _
        " As Credit_In"

cField1 = cField1 & "," & _
         myiif( _
        " (FLAG = 203)", "PLUS") & _
        " As Trust_cash_In"


' ----------- ’«œ—
cField2 = myiif( _
        " (FLAG = 2 or flag = 3)", "MINUS - PLUS") & _
        " As Cash_out"

cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 111 or FLAG = 112)", "MINUS") & _
        " As Pur_cash"

cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 15)", "MINUS") & _
        " As BANK_OUT"

cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 7)", "MINUS") & _
        " As TRANS_OUT"

cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 5 OR FLAG = 101)", "MINUS") & _
        " As CHARGES"
        
cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 16)", "MINUS") & _
        " As CHQ_OUT"
        
cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 18)", "MINUS") & _
        " As PART_OUT"
        
cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 20)", "MINUS") & _
        " As DEBIT_OUT"

cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 22)", "MINUS") & _
        " As CREDIT_OUT"
        
cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 102)", "MINUS") & _
        " As Trust_Out"
        
cField2 = cField2 & "," & _
         myiif( _
        " (FLAG = 202)", "MINUS") & _
        " As Trust_cash_Out"

cwhere = ""
If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If XCODE.Text <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XCODE.Text)
End If

cField3 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as Last_Balance"

cString = "Select " & cField1 & "," & cField2 & "," & cField3 & _
           " From BOXMOVE "

cwhere = ""
If IsDate(xdate1.Text) Then cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)

If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If Trim(XCODE.Text) <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " BOX = " & MyParn(XCODE.Text)
End If

cString = cString & turn(cwhere) & cwhere
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (sourcetable.EOF And sourcetable.BOF) Then
    xFirst_Balance.Caption = Myvalue(Round(Val(sourcetable!First_Balance & "") + Val(sourcetable!First_Bal & ""), 2))
    xCash_in.Caption = Myvalue(sourcetable!Cash_in)
    xChq_in.Caption = Myvalue(sourcetable!Chq_in)
    xBank_in.Caption = Myvalue(sourcetable!Bank_in)
    xTrans_in.Caption = Myvalue(sourcetable!Trans_In)
    xIncome.Caption = Myvalue(sourcetable!Income)
    xTrust_in.Caption = Myvalue(sourcetable!trust_in)
    xTrust_cash_in.Caption = Myvalue(sourcetable!trust_cash_in)
    xDebit_In.Caption = Myvalue(sourcetable!debit_in)
    xPart_In.Caption = Myvalue(sourcetable!Part_in)
    xCredit_in.Caption = Myvalue(sourcetable!Credit_in)
    xTotal_in.Caption = Val(xFirst_Balance.Caption) + Val(xCash_in.Caption) + Val(Me.xChq_in.Caption) + _
                        Val(xBank_in.Caption) + Val(xTrans_in.Caption) + Val(xIncome.Caption) + _
                        Val(xTrust_in.Caption) + Val(xTrust_cash_in.Caption) + Val(xDebit_In.Caption) + _
                        Val(xPart_In.Caption) + Val(xCredit_in.Caption)
    
    xCash_out.Caption = Myvalue(sourcetable!Cash_out)
    xPur_cash.Caption = Myvalue(sourcetable!Pur_Cash)
    xBank_out.Caption = Myvalue(sourcetable!Bank_Out)
    xTrans_out.Caption = Myvalue(sourcetable!Trans_Out)
    xCharges.Caption = Myvalue(sourcetable!CHARGES)
    xChq_Out.Caption = Myvalue(sourcetable!Chq_Out)
    xPart_Out.Caption = Myvalue(sourcetable!Part_Out)
    xDebit_Out.Caption = Myvalue(sourcetable!Debit_Out)
    xTrust_out.Caption = Myvalue(sourcetable!trust_out)
    xTrust_cash_out.Caption = Myvalue(sourcetable!trust_cash_Out)
    xTotal_out.Caption = Val(xCash_out.Caption) + Val(xBank_out.Caption) + Val(xTrans_out) + _
            Val(xCharges.Caption) + Val(xChq_Out.Caption) + Val(xPart_Out.Caption) + _
            Val(xTrust_out.Caption) + Val(xTrust_cash_out.Caption) + Val(xDebit_Out.Caption) _
            + Val(xCredit_out.Caption) + Val(xPur_cash.Caption)
    xLast_Balance.Caption = Myvalue(Val(sourcetable!Last_Balance & ""))
End If
sourcetable.Close
Set sourcetable = Nothing
End Sub
Sub myProc()
ActiveControl.Text = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
Unload oSearchBox
End Sub
Function MYVALID() As Boolean
If XCODE.Text = "" Then
    MsgBox "ﬂÊœ «·Œ“‰… €Ì— „”Ã·"
    Exit Function
End If
If IsEmpty(GetField("select Desca from file0_50 where code = " & MyParn(XCODE.Text), con)) Then
    MsgBox "ﬂÊœ «·Œ“‰… €Ì— ’ÕÌÕ"
    Exit Function
End If
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(XDATE2.Text)) And Trim(XDATE2.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
myloadgrd
myloadgrd2
MyLoadTotal
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
cHeader1 = "Õ—ﬂ… Œ“Ì‰… " & xDesca.Caption
If IsDate(xdate1.Text) Then cHeader2 = BetweenString(xdate1.Text, XDATE2.Text)
If IsDate(XDATE2.Text) Then cHeader2 = BetweenString(xdate1.Text, XDATE2.Text)

Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", grid1.Rows - 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)

PrintGrdNew.doprint grid1, 0.9, -1, cHeader1, cHeader2, , , False, False, 10, , aRow
PrintGrdNew.Show 1
'MyAddItem
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdPrint1_Click()
doPrintTotal
End Sub

Private Sub CmdPrint2_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
cHeader1 = "≈Ã„«·Ì Õ—ﬂ… ÌÊ„Ì… ·Œ“Ì‰… " & xDesca.Caption
If IsDate(xdate1.Text) Then cHeader2 = BetweenString(xdate1.Text, XDATE2.Text)
If IsDate(XDATE2.Text) Then cHeader2 = BetweenString(xdate1.Text, XDATE2.Text)
PrintGrdNew.doprint grid2, 1, -1, cHeader1, cHeader2, , , False, False, 10
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
openCon con
Fixgrd
fixgrd2
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
On Error Resume Next
Unload Search3
Err.Clear
End Sub

Private Sub Grid1_DblClick()
    Dim cDoc_no As String
    Select Case grid1.TextMatrix(grid1.Row, 6)
        Case "4", "5"
            cDoc_no = grid1.TextMatrix(grid1.Row, 2)
            salesfrm.myPublic = IIf(grid1.TextMatrix(grid1.Row, 6) = "4", 0, 1)
            salesfrm.sDoc_no = cDoc_no
            salesfrm.Show 1
        Case "A", "C"
    End Select
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xCode_Change()
grid1.Rows = 1
cmdGo.Enabled = Trim(XCODE.Text) <> ""
End Sub
Private Sub XCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then BoxLookupAll Me, oSearchBox
End Sub
Private Sub xCode_LostFocus()
xDesca.Caption = ""
If Trim(XCODE.Text) = "" Then Exit Sub
XCODE.Text = RetZero(XCODE.Text, 6)
xDesca.Caption = GetField("select Desca from file0_50 where code = " & MyParn(XCODE.Text), con) & ""
End Sub
Private Sub xStore_Click(Area As Integer)
cmdGo.Enabled = True
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file3_10"
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
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub doprint()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim n11 As Double, n12 As Double, n13 As Double, n14 As Double, n15 As Double, n16 As Double, n17 As Double
Dim cStrW As String

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If Trim(XCODE.Text) <> "" Then
    aHeader(0) = "[" & "··⁄„Ì· : " & xDesca.Caption & "]"
End If
If IsDate(xdate1.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
    cStrW = cStrW & " AND DATE >= " & DateSq(xdate1.Text)
End If
If IsDate(XDATE2.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
    cStrW = cStrW & " AND DATE <= " & DateSq(XDATE2.Text)
End If

n11 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE3_11 WHERE [TYPE] = '4' AND CODE = " & MyParn(XCODE.Text) & cStrW) & "")
n12 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '5' AND CODE = " & MyParn(XCODE.Text) & cStrW) & "")
n13 = n11 - n12
n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = '10' OR [TYPE] = '11' OR [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(XCODE.Text) & cStrW) & "")


n15 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = 'A' OR [TYPE] = 'C' ) AND CODE = " & MyParn(XCODE.Text) & cStrW) & "")
n16 = n14 + n15
n17 = Val(GetDesca("SELECT SUM(VALUE)  FROM FILE5_20 WHERE [CLOSED] = '0' AND CODE1 = " & MyParn(XCODE.Text)) & "")

With grid1
For i = 1 To .Rows - 2
    temptable.AddNew
    temptable!Date1 = TurnValue(RealDate(.TextMatrix(i, 1)))
    temptable!str1 = TurnValue(.TextMatrix(i, 2))
    temptable!str2 = TurnValue(.TextMatrix(i, 0))
    temptable!val1 = Val(.TextMatrix(i, 3))
    temptable!val2 = Val(.TextMatrix(i, 4))
    temptable!val3 = Val(.TextMatrix(i, 5))
    temptable!Val6 = i
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!STR20 = Firsttitle
    
    temptable!VAL10 = n11
    temptable!VAL11 = n12
    temptable!val12 = n13
    temptable!val13 = n14
    temptable!VAL14 = n15
    temptable!VAL15 = n16
    temptable!Val16 = n17
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\client3.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Function RealDate(pDate, Optional cFormat As String = "") As String
If Not IsDate(pDate) Then Exit Function
RealDate = DateValue(Format(pDate, "dd/mm/yyyy"))
If cFormat <> "" Then RealDate = Format(RealDate, cFormat)
End Function
Private Sub myloadgrd2()
Dim sourcetable As New ADODB.Recordset, nBalance As Double, cwhere As String

If Trim(XCODE.Text) <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " BOX = " & MyParn(XCODE.Text)
End If

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
    cField1 = "(" & _
               "Select Sum(PLUS - MINUS) From BoxMove " & _
               turn(cwhere) & cwhere & _
               ") as FirstBalance"
Else
    cField1 = "0 AS FirstBalance"
End If


cString = "Select Date, " & cField1 & ", Sum(PLUS) as sumofPlus,sum(Minus) as sumofMinus,Sum(Plus- MINUS) as SumOfValue " & _
           " From boxmove "

cwhere = ""
If IsDate(xdate1.Text) Then cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
If XCODE.Text <> "" Then cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XCODE.Text)
cString = cString & turn(cwhere) & cwhere & " Group by Date ORDER BY DATE"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With grid2
.Rows = 1
If Not (sourcetable.EOF And sourcetable.EOF) Then
    nBalance = Val(sourcetable!FirstBalance & "")
    If Val(nBalance) <> 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "—’Ìœ ”«»ﬁ"
        .TextMatrix(.Rows - 1, 2) = Myvalue(nBalance)
    End If
End If

Do Until sourcetable.EOF
    If Val(sourcetable!Sumofvalue & "") <> 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = Format(sourcetable!Date, "YYYY/MM/DD")
        .TextMatrix(.Rows - 1, 1) = Myvalue(sourcetable!SumofPlus)
        .TextMatrix(.Rows - 1, 2) = Myvalue(sourcetable!SumofMinus)
        .TextMatrix(.Rows - 1, 3) = Myvalue(sourcetable!Sumofvalue)
        nBalance = Val(sourcetable!Sumofvalue & "") + nBalance
        .TextMatrix(.Rows - 1, 4) = Myvalue(nBalance)
    End If
    sourcetable.MoveNext
Loop

    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 1, "#", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 2, "#", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 3, "#", vbYellow, vbRed, True, "  "
    If grid1.Rows > 1 Then
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 4) = .TextMatrix(.Rows - 1, 3)
    End If

End With
sourcetable.Close
Set sourcetable = Nothing
End Sub
Private Sub fixgrd2()
With grid2
    .Rows = 1
    .TextMatrix(0, 0) = "«· «—ÌŒ"
    .TextMatrix(0, 1) = "≈Ã„«·Ì Ê«—œ"
    .TextMatrix(0, 2) = "≈Ã„«·Ì ’«œ—"
    .TextMatrix(0, 3) = "«·’«›Ì"
    .TextMatrix(0, 4) = "—’Ìœ «·ÌÊ„"
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
End With
End Sub
Private Sub Fixgrd()
With grid1
.TextMatrix(0, 0) = "‰Ê⁄ «·Õ—ﬂ…"
.TextMatrix(0, 1) = "«·»Ì«‰"
.TextMatrix(0, 2) = " «—ÌŒ"
.TextMatrix(0, 3) = "„” ‰œ"
.TextMatrix(0, 4) = "„œÌ‰"
.TextMatrix(0, 5) = "œ«∆‰"
.TextMatrix(0, 6) = "—’Ìœ"

.ColWidth(0) = 3000
.ColWidth(1) = 2000
.ColWidth(2) = 1400
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub doPrintTotal()
Dim temptable As New ADODB.Recordset, aHeader(1)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
contemp.BeginTrans
If IsDate(xdate1.Text) Or IsDate(XDATE2.Text) Then
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If
If Trim(XCODE.Text) <> "" Then
    aHeader(0) = "≈Ã„«·Ì —’Ìœ Œ“‰… : " & xDesca.Caption
End If

temptable.AddNew
temptable!val1 = Val(xFirst_Balance.Caption)
temptable!val2 = Val(xCash_in.Caption)
temptable!val3 = Val(xBank_in.Caption)
temptable!val4 = Val(xTrans_in.Caption)
temptable!VAL5 = Val(xChq_in.Caption)
temptable!Val6 = Val(xIncome.Caption)
temptable!VAL7 = Val(xTrust_in.Caption)
temptable!VAL8 = Val(xTrust_cash_in.Caption)
temptable!VAL9 = Val(xDebit_In.Caption)
temptable!VAL10 = Val(xPart_In.Caption)

temptable!VAL11 = Val(xCash_out.Caption)
temptable!val12 = Val(xBank_out.Caption)
temptable!val13 = Val(xTrans_out.Caption)
temptable!VAL14 = Val(xChq_Out.Caption)
temptable!VAL15 = Val(xCharges.Caption)
temptable!Val16 = Val(xTrust_out.Caption)
temptable!Val16 = Val(xTrust_cash_out.Caption)
temptable!Val17 = Val(xDebit_Out.Caption)
temptable!Val19 = Val(xPart_Out.Caption)

temptable!Val23 = Val(xTotal_in.Caption)
temptable!Val24 = Val(xTotal_out.Caption)
temptable!Val25 = Val(xLast_Balance.Caption)
temptable!str11 = TurnValue(retHeader(aHeader, 0, 1))
temptable!str12 = TurnValue(retHeader(aHeader, 1, 1))
temptable.Update
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\BALBOX.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

