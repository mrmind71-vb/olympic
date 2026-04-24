VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form arrivefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16020
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
   ScaleHeight     =   9495
   ScaleWidth      =   16020
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   1185
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   2085
      Begin VB.CommandButton Command3 
         Caption         =   "«·€«¡ „” ‰œ «·„’—Ê›"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   630
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   " ÕÊÌ· «·Ì „’—Ê›"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   135
         Width           =   1995
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1485
      Width           =   2265
      Begin VB.CommandButton cmdAddEmp 
         Caption         =   "«÷«›… «·„ÊŸ›Ì‰"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1035
      Width           =   1275
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "arrive.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   585
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "arrive.frx":2579
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "arrive.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "arrive.frx":6CFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "arrive.frx":9594
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "arrive.frx":BB40
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   1545
      Begin VB.CommandButton CMDPRINT 
         Height          =   510
         Left            =   45
         Picture         =   "arrive.frx":E313
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "√÷«›… »‰›” «·»Ì«‰« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1410
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   9465
      Begin VB.TextBox xcharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3195
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox xWeek 
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
         Height          =   375
         Left            =   3195
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox xMonth 
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
         Height          =   375
         Left            =   6750
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   510
      End
      Begin VB.TextBox xYear 
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
         Height          =   375
         Left            =   8055
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   690
      End
      Begin VB.TextBox xdate 
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
         Height          =   375
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1680
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
         Height          =   360
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   945
         Width           =   8655
      End
      Begin VB.TextBox xdoc_no 
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
         Height          =   375
         Left            =   -495
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   315
         Visible         =   0   'False
         Width           =   960
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   330
         Left            =   5625
         TabIndex        =   37
         Top             =   585
         Width           =   3120
         _ExtentX        =   5503
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
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„’—Ê›"
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   630
         Width           =   795
      End
      Begin VB.Label xCharge_Desca 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   585
         Width           =   3075
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“‰…"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label12 
         Caption         =   "«·«”»Ê⁄"
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   255
         Width           =   660
      End
      Begin VB.Label xMonthDesca 
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
         Height          =   375
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Height          =   240
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "«·‘Â—"
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
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "«·”‰…"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   570
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5940
      Top             =   135
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Left            =   3555
      Top             =   585
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
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
      Left            =   945
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6270
      Left            =   90
      TabIndex        =   26
      Top             =   2160
      Width           =   15810
      _cx             =   27887
      _cy             =   11060
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
      Cols            =   11
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   8415
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   29
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
         Picture         =   "arrive.frx":1073D
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "arrive.frx":1290D
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   30
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
         Picture         =   "arrive.frx":14A55
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "arrive.frx":16C1D
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   31
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
         Picture         =   "arrive.frx":18D6C
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "arrive.frx":1AF4C
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   32
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
         Picture         =   "arrive.frx":1D0A7
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "arrive.frx":1F263
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8415
      Visible         =   0   'False
      Width           =   3165
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Visible         =   0   'False
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "arrivefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bedit As Boolean
Dim formMode, dDateLast As String
Dim oSearch As New Search3, oSearchEmp As New Search3, oSearchCharge As New Search3
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[date]", addDate(xdate.Text))
aInsert = AddFlag(aInsert, "[YEAR]", Val(xYear.Text))
aInsert = AddFlag(aInsert, "[MONTH]", Val(xMonth.Text))
aInsert = AddFlag(aInsert, "[WEEK]", Val(xWeek.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "[BOX]", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "[CHARGE]", addstring(xcharge.Text))
aInsert = AddFlag(aInsert, "[TYPE]", "2")

con.BeginTrans
On Error GoTo myerror
If xdoc_no.Text = "" Then
    xdoc_no.Text = Newflag("FILE2_50H", "doc_no")
    aInsert = AddFlag(aInsert, "[DOC_NO]", addvalue(xdoc_no.Text))
    con.Execute addInsert(aInsert, "FILE2_50H")
Else
    con.Execute addUpdate(aInsert, "FILE2_50H", "doc_no = " & xdoc_no.Text)
End If
myReplacegrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = FoundOtherEdit(grid1.Row, 1, oSearchEmp.grid1.TextMatrix(oSearchEmp.grid1.Row, 0))
    If nFound <> -1 Then
        MsgBox ("«·„ÊŸ› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound)
        Exit Sub
    End If
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearchEmp.grid1.TextMatrix(oSearchEmp.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = oSearchEmp.grid1.TextMatrix(oSearchEmp.grid1.Row, 1)
    grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearchEmp
        CellPos 13, grid1.Row, 2
    Else
        grid1.Select grid1.Rows - 1, 1
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xdoc_no.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
    myUndo
ElseIf ActiveControl.Name = xcharge.Name Then
    xcharge.Text = oSearchCharge.grid1.TextMatrix(oSearchCharge.grid1.Row, 0)
    xCharge_Desca.Caption = oSearchCharge.grid1.TextMatrix(oSearchCharge.grid1.Row, 1)
    Unload oSearchCharge
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE2_50 where Doc_No = " & MyParn(xdoc_no.Text)
    con.Execute "Delete  From FILE2_50H where Doc_No = " & MyParn(xdoc_no.Text)
    con.CommitTrans
    openCardTable
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
       CardTable.Find "Doc_No < " & MyParn(xdoc_no.Text), , adSearchBackward, adBookmarkLast
       If CardTable.BOF Then CardTable.MoveFirst
       myload
    End If
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
'If Not bNoMsgExit Then If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
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
Private Sub CmdNewInv_Click()
mydefine
On Error Resume Next
xMonth.SetFocus
Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdPrint_Click()
Dim cHeader1 As String
cHeader1 = "Õ÷Ê— «·ÌÊ„Ì… ⁄‰ ‘Â— " & xMonthDesca.Caption & Space(1) & "”‰… " & xYear.Text
grid1.Rows = grid1.Rows - 1
PrintGrdNew.doprint grid1, 0.8, -3, cHeader1, , , , False, False, 10
PrintGrdNew.Show 1
MyAddItem
End Sub
Private Sub Command1_Click()
xdoc_no.Text = ""
xdate.Text = Format(Date, "dd-mm-yyyy")
xMonth.Text = ""
xYear.Text = Year(Date)
xDesca.Text = ""
With grid1
For I = 1 To grid1.Rows - 2
    grid1.TextMatrix(I, .Cols - 1) = ""
Next
End With
Handlecontrols DefineMode
End Sub

Private Sub Command3_Click()
'mySave False
'doprint App.Path & "\Reports\salary2.rpt"
doprint
End Sub
Private Sub cmdAddEmp_Click()
Dim loctable As New ADODB.Recordset, nRowAdd As Long
If Not (IsNumeric(xMonth.Text) And IsNumeric(xYear.Text) And IsDate(xdate.Text)) Then
    MsgBox "«·‘Â— Ê«·”‰… «Ê«· «—ÌŒ €Ì— „”Ã·Ì‰"
    Exit Sub
End If

If Not (Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12) Then
    MsgBox "«·‘Â— €Ì— ’ÕÌÕ"
    Exit Sub
End If

cString = "select code,desca,Salary  from driver where salary_code = 2"
cString = cString & turn(cString) & "((date_end is null) or Date_end >= " & DateSq(xdate.Text) & ")"
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If grid1.FindRow(loctable!code & "", , 1) = -1 Then
        grid1.TextMatrix(grid1.Rows - 1, 1) = loctable!code
        grid1.TextMatrix(grid1.Rows - 1, 2) = loctable!desca
        grid1.TextMatrix(grid1.Rows - 1, 3) = loctable!Salary
        nRowAdd = nRowAdd + 1
        grid1.AddItem ""
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
If nRowAdd = 0 Then Exit Sub
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub

Private Sub Form_Activate()
On Error Resume Next
If xdoc_no.Tag = LoadMode Then
    grid1.SetFocus
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
bedit = True
openCon con

Set data1.Recordset = myRecordSet("Select * From file0_50 where code > '500000'", con)
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Err.Clear
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    Calctotals
    Exit Sub
End If
With grid1
If Row = grid1.Rows - 1 Then
    MyAddItem
End If
Calctotals
If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        xdoc_no.Tag = LoadMode
        xdoc_no.Enabled = False
    End If
    If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadgrd
    End If
End If
End With
End Sub
Private Sub grid1_EnterCell()
If (grid1.Col = 1 Or grid1.Col = 4 Or grid1.Col = 6 Or grid1.Col = 7 Or grid1.Col = 9) Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.Select 1, 1
End If
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Val(xYear.Text) < 2010 Or Val(xYear.Text) > 2030 Or Not IsNumeric(xYear.Text) Then
    If Not bIgMsg Then MsgBox "«·”‰… €Ì— ”·Ì„…"
    Exit Function
End If
If Val(xWeek.Text) < 1 Or Val(xWeek) > 4 Then
    If Not bIgMsg Then MsgBox "«·«”»Ê⁄ €Ì— ”·Ì„"
    Exit Function
End If

If Val(xMonth.Text) < 1 Or Val(xMonth.Text) > 12 Then
    If Not bIgMsg Then MsgBox "«·‘Â— €Ì— ”·Ì„"
    Exit Function
End If

If Not IsDate(xdate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload()
xdoc_no.Text = CardTable!doc_no
xdate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xMonth.Text = CardTable!Month & ""
xWeek.Text = CardTable!WEEK & ""
xMonthDesca.Caption = arbMonth(Val(CardTable!Month & ""))
xYear.Text = CardTable!Year & ""
xDesca.Text = CardTable!desca & ""
xBox.BoundText = CardTable!BOX & ""
xcharge.Text = CardTable!CHARGE & ""
xCharge_Desca.Caption = GetField("select desca from file8_51 where code = " & MyParn(xcharge.Text)) & ""
myloadgrd
Handlecontrols LoadMode
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub mydefine()
xdoc_no.Text = ""
xdate.Text = ""
xMonthDesca.Caption = ""
xMonth.Text = ""
xYear.Text = Year(Date)
xDesca.Text = ""
xBox.BoundText = ""
xWeek.Text = ""
xcharge.Text = GetField("SELECT CHARGE FROM FILE2_50H WHERE (NOT CHARGE IS NULL) AND TYPE = 2 ORDER BY DOC_NO DESC") & ""
If xcharge.Text <> "" Then xCharge_Desca.Caption = GetField("select desca from file8_51 where code = " & MyParn(xcharge.Text)) & ""
grid1.Rows = 1
MyAddItem
Fixgrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
CMDPRINT.Enabled = (nMode = LoadMode)
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And doc_no = ""
xMonth.Enabled = (nMode = DefineMode)
xYear.Enabled = (nMode = DefineMode)
xWeek.Enabled = (nMode = DefineMode)
xdoc_no.Enabled = (nMode = DefineMode)
xdoc_no.Tag = nMode
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim nFound As Long
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ  €Ì— „”Ã·"
        Cancel = True
    Else
        grid1.EditText = RetZero(grid1.EditText)
        nFound = FoundOtherEdit(Row, Col, RetZero(grid1.EditText))
        If nFound <> -1 Then
            MsgBox "«·„ÊŸ› „ﬂ—— ›Ï «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
        Else
            grid1.EditText = RetZero(grid1.EditText)
            aRet = GetFields("select * from driver where SALARY_CODE = 2 and code = " & MyParn(grid1.EditText), con)
            If IsEmpty(aRet) Then
               MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
               Cancel = True
            Else
                grid1.TextMatrix(Row, 2) = retFlag(aRet, "desca") & ""
                grid1.TextMatrix(Row, 3) = retFlag(aRet, "salary") & ""
            End If
        End If
    End If
End If
End Sub

Private Sub XCHARGE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ChargeLookupAll Me, oSearchCharge
End Sub

Private Sub xCharge_Validate(Cancel As Boolean)
xCharge_Desca.Caption = ""
If xcharge.Text = "" Then Exit Sub
xcharge.Text = RetZero(xcharge.Text, 3)
Dim aRet As Variant
aRet = GetFields("select code,desca from file8_51 where code = " & MyParn(xcharge.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·„’—Ê› €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xCharge_Desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub

Private Sub XDATE_DblClick()
Set datefrm.oDate = xdate
datefrm.Show 1
End Sub
Private Sub xDoc_No_LostFocus()
If xdoc_no.Text = "" Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_no = " & MyParn(xdoc_no.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xdoc_no.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "Delete from  FILE2_50 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        myRemove grid1.Row
    End If
ElseIf KeyCode = 112 Then
    DriverLookupAll Me, oSearchEmp, "SALARY_CODE = 2"
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub GrdDesc(Row)
'If Not ValidInt(grid1.TextMatrix(Row, 0)) Then Exit Sub
'grid1.TextMatrix(Row, 1) = GetDesca("Select desca from DRIVER where code = " & grid1.TextMatrix(Row, 0))
End Sub
Private Function Calctotals()
Dim nTotal As Single, nTotal2 As Single, nTotal3 As Single
With grid1
For I = 1 To grid1.Rows - 2
    grid1.TextMatrix(I, 5) = Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4))
    grid1.TextMatrix(I, 8) = Val(grid1.TextMatrix(I, 5)) + Val(grid1.TextMatrix(I, 6)) - Val(grid1.TextMatrix(I, 7))
'    nTotal1 = nTotal1 + Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7))
'    grid1.TextMatrix(i, 12) = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 10)) - Val(.TextMatrix(i, 11))
'    nTotal2 = nTotal2 + Val(.TextMatrix(i, 9)) + Val(.TextMatrix(i, 10)) + Val(.TextMatrix(i, 11))
'    nTotal = nTotal + Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 10)) - Val(.TextMatrix(i, 11))
Next
'xTotal1.Caption = Format(nTotal1, "Fixed")
'xTotal2.Caption = Format(nTotal2, "Fixed")
'xtotal.Caption = Format(nTotal, "Fixed")
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO, CONVERT(VARCHAR(1),[WEEK]), CONVERT(NVARCHAR(2),[Month]), CONVERT(NVARCHAR(4),[Year]) ,convert(VARCHAR(10),[DATE],111),Desca " & _
                  " FROM  FILE2_50H WHERE TYPE  = 1"

Generalarray(2) = "Order by Year,Month,DOC_NO "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·”‰…"
listarray(0, 1) = "(**[year]**)"

listarray(0, 0) = "«·‘Â—"
listarray(0, 1) = "(**[MONTH]**)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”»Ê⁄"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«·‘Â—"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "«·”‰…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "«· «—ÌŒ"
GrdArray(4, 1) = 1400

GrdArray(5, 0) = "»Ì«‰"
GrdArray(5, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„"
oSearch.Show 1
End Sub
Private Function foundOther(Optional ByVal Col As Long, Optional ByVal Row As Long = -1) As Integer
Dim sfind As String
foundOther = -1
For I = 1 To grid1.Rows - 1
    If I <> Row And Trim(grid1.TextMatrix(I, nCol)) <> "" Then
        If LCase(Trim(grid1.TextMatrix(I, Col))) = LCase(Trim(grid1.TextMatrix(I, Col))) Then
            foundOther = I
            Exit Function
        End If
    End If
Next
End Function
Private Function FoundOtherEdit(ByVal Row, ByVal Col As Long, ByVal sfind As String) As Long
FoundOtherEdit = -1
If Trim(LCase(grid1.TextMatrix(Row, Col))) = Trim(LCase(sfind)) Then
    Exit Function
End If

For I = 1 To grid1.Rows - 1
    If I <> Row And Trim(grid1.TextMatrix(I, Col)) <> "" Then
        If LCase(Trim(sfind)) = LCase(Trim(grid1.TextMatrix(I, Col))) Then
            FoundOtherEdit = I
           Exit Function
        End If
    End If
Next
End Function
Private Sub Fixgrd()
With grid1
.MergeRow(0) = True
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.FormatString = "„|" & "ﬂÊœ|" & "«·«”„|" & "«·ÌÊ„Ì…|" & "«Ì«„ «·Õ÷Ê—|" & "«·„— »|" & "«÷«›Ì|" & "Ã“«¡« |" & "«·«Ã„«·Ì|" & "„·«ÕŸ« |"
.ColWidth(0) = 500
.ColWidth(1) = 800
.ColWidth(2) = 3000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 3000

.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub doprint(Optional nFlag As Long = 0, Optional cEmpCode As String = "")
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    If cEmpCode = "" Or (Trim(grid1.TextMatrix(I, 0)) = cEmpCode) Then
        temptable.AddNew
        temptable!str21 = "Õ÷Ê— «·ÌÊ„Ì… ⁄‰ ‘Â— " & xMonthDesca.Caption & Space(1) & "”‰… " & xYear.Text
        temptable!str1 = TurnValue(grid1.TextMatrix(I, 1))
        
        temptable!val1 = Val(grid1.TextMatrix(I, 2))
        temptable!val2 = Val(grid1.TextMatrix(I, 3))
        
        temptable!Val3 = Val(grid1.TextMatrix(I, 4))
        
        temptable!val5 = Val(grid1.TextMatrix(I, 5))
        temptable!Val6 = Val(grid1.TextMatrix(I, 6))
        temptable!Val7 = Val(grid1.TextMatrix(I, 7))
        
        temptable!Val8 = Val(grid1.TextMatrix(I, 8))
        temptable!val9 = Val(grid1.TextMatrix(I, 9))
        temptable!Val10 = Val(grid1.TextMatrix(I, 10))
        temptable!val11 = Val(grid1.TextMatrix(I, 11))
        temptable!val12 = Val(grid1.TextMatrix(I, 12))
        temptable!str2 = TurnValue(grid1.TextMatrix(I, 13))
         
        temptable!Str3 = TurnValue(xDesca_w1.Text)
        temptable!str4 = TurnValue(xDesca_w2.Text)
        temptable!str5 = TurnValue(xDesca_w3.Text)
        temptable!str6 = TurnValue(xDesca_w4.Text)
        temptable!Val20 = nFlag
        
        temptable.Update
    End If
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\salary2.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function myReplacegrd(Row As Long) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xdoc_no.Text))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "SALARY", Val(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DAYS", Val(grid1.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "PLUS", Val(grid1.TextMatrix(I, 6)))
        aInsert = AddFlag(aInsert, "MINUS", Val(grid1.TextMatrix(I, 7)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(I, 9)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE2_50")
        Else
            con.Execute addUpdate(aInsert, "FILE2_50", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Function
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE2_50.CODE,DRIVER.DESCA,FILE2_50.SALARY, FILE2_50.DAYS,FILE2_50.SALARY_TOTAL,FILE2_50.PLUS,FILE2_50.MINUS,FILE2_50.TOTAL, FILE2_50.NOTES,FILE2_50.ID" & _
              " FROM FILE2_50 inner JOIN DRIVER ON FILE2_50.CODE = DRIVER.CODE WHERE DOC_NO = " & MyParn(xdoc_no.Text)
    DATA11.RecordSource = cString
    DATA11.Refresh
    grid1.AddItem ""
End With
Calctotals
Fixgrd
MakeSerial
End Sub
Private Sub xMonth_Validate(Cancel As Boolean)
If (Not ValidInt(xMonth.Text)) Then
    Cancel = True
ElseIf Not (xMonth.Text >= 1 And xMonth.Text <= 12) Then
    MsgBox "«·‘Â— ·« Ì’·Õ"
    Cancel = True
Else
    xMonthDesca.Caption = arbMonth(xMonth.Text)
    LoadData
End If
End Sub

Private Sub xWeek_Validate(Cancel As Boolean)
If (Not ValidInt(xWeek.Text)) Then
    Cancel = True
ElseIf Not (xWeek.Text >= 1 And xWeek.Text <= 4) Then
    MsgBox "«·«”»Ê⁄ ·« Ì’·Õ"
    Cancel = True
Else
    LoadData
End If
End Sub

Private Sub xYear_Validate(Cancel As Boolean)
If (Not ValidInt(xYear.Text)) Then
    Cancel = True
ElseIf Not (Val(xYear.Text) >= 2014 And Val(xYear.Text) <= 2040) Then
    MsgBox "«·”‰… ·«  ’·Õ"
    Cancel = True
Else
    LoadData
End If
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM FILE2_50H WHERE TYPE = 2"
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
cString = cString & " Order by FILE2_50H.DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xdoc_no.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xdoc_no.Text), , adSearchForward, adBookmarkFirst
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
.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
End With
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then myRemove OldRow
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
If Not MYVALID(True) Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1 + IIf(Col = 4 Or Col = 7, 1, 0) + IIf(Col = 1, 2, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 4)
    grid1.ShowCell grid1.Row, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial
Calctotals
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Function LoadData() As Boolean
If Not validData Then
    If xdoc_no.Tag = LoadMode Then mydefine
    Exit Function
End If

Dim sDoc_no As Variant
sDoc_no = GetField("Select doc_no from FILE2_50H where [YEAR] = " & xYear.Text & _
          " and [Month] = " & xMonth.Text & " AND WEEK = " & xWeek.Text & " AND TYPE = 2")
If Not IsEmpty(sDoc_no) Then
    xdoc_no.Text = sDoc_no
    myUndo
ElseIf xdoc_no.Tag = LoadMode Then
    mydefine
End If
End Function
Private Function validData() As Boolean
If Not ((ValidInt(xMonth.Text)) And ValidInt(xYear.Text) And ValidInt(xWeek.Text)) Then Exit Function
If Not (Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12) Then Exit Function
If Not (Val(xYear.Text) >= 2014 And Val(xYear.Text) <= 2040) Then Exit Function
If Not (Val(xWeek.Text) >= 1 And Val(xWeek.Text) <= 4) Then Exit Function
validData = True
End Function

