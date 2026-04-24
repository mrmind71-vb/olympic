VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form trustfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19005
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
   ScaleHeight     =   9090
   ScaleWidth      =   19005
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   8280
      Width           =   10140
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " „  ”ÊÌ Â«"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "·„ Ì „  ”ÊÌ Â«"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ﬂ· «·„” ‰œ« "
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   8505
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAddTravel 
      Caption         =   "«÷«›… «·⁄„·Ì« "
      Height          =   465
      Left            =   8910
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1530
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   11745
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "trust2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust2.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust2.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust2.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust2.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   180
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
         Height          =   510
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust2.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "«· ”ÊÌ…"
      Height          =   5775
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1980
      Width           =   7665
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5415
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   7440
         _cx             =   13123
         _cy             =   9551
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
         Cols            =   5
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
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8280
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   13
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
         Picture         =   "trust2.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust2.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   14
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
         Picture         =   "trust2.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust2.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   15
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
         Picture         =   "trust2.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust2.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   16
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
         Picture         =   "trust2.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust2.frx":1CE87
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1755
      Top             =   -135
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -2610
      Top             =   90
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   -1215
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   810
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   3240
      Top             =   -135
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -3105
      Top             =   -135
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
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "≈€·«ﬁ «·„” ‰œ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9045
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame Frame2 
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
      Left            =   10620
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   675
      Width           =   8340
      Begin VB.CheckBox xDone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " „  «· ”ÊÌ…"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox xTrust_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2925
         MaxLength       =   300
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   900
         Width           =   4200
      End
      Begin VB.TextBox xBox 
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
         Left            =   6030
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   345
         Left            =   6030
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox xDate 
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
         Height          =   345
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ”ÊÌ…"
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
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   900
         Width           =   870
      End
      Begin VB.Label xBox_desca 
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   3075
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   885
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰…"
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
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   -1215
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
   Begin VB.Frame Frame4 
      Caption         =   "„” ‰œ«  «· ”ÊÌ…"
      Height          =   5730
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1980
      Width           =   11220
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5370
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   11040
         _cx             =   19473
         _cy             =   9472
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
         Cols            =   7
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
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   8865
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7650
      Width           =   10095
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label xTotal_Desca 
         Caption         =   "«·»«ﬁÌ"
         Height          =   240
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   600
      End
      Begin VB.Label xTotal_Cost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label xTotal_trust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "≈Ã„«·Ì «·⁄Âœ…"
         Height          =   330
         Left            =   8865
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "≈Ã„«·Ì «·„’—Ê›"
         Height          =   285
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   7650
      Width           =   8745
      Begin VB.Label xTotal_M 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "”Õ»"
         Height          =   285
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label13 
         Caption         =   "—’Ìœ"
         Height          =   330
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label12 
         Caption         =   "«Ìœ«⁄ "
         Height          =   285
         Left            =   7965
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   270
         Width           =   645
      End
      Begin VB.Label xTotal_P 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   1500
      End
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      Caption         =   " „   ”ÊÌ… «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   510
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1530
      Width           =   3120
   End
End
Attribute VB_Name = "trustfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, bSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String
Dim cFilter As String, cList As String, clist1 As String, cList2 As String, cPlaceString As String
Dim oSearchDoc As New Search3, osearchitem As New Search3, oSearchBox As New Search3, oSearch_Travel As New Search3, oSearchSup As New Search3, oSearchDriver As New Search3
Dim oAdd As New trust_addfrm
Dim bEdit As Boolean
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "BOX", addstring(XBOX.Text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[TRUST_NO]", addstring(xTrust_No.Text))
aInsert = AddFlag(aInsert, "[DONE]", xDone.Value)
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Text = "" Then
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(sUserName))
    xDoc_No.Text = RetZero(Val(Newflag("TRUST_H", "doc_no")))
    aInsert = AddFlag(aInsert, "[DOC_NO]", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "TRUST_H")
Else
    con.Execute addUpdate(aInsert, "TRUST_H", "doc_no = " & addstring(xDoc_No.Text))
End If
If Row <> -1 And Row2 = -1 Or (Row = -1 And Row2 = -1) Then myreplaceGrd Row
If Row = -1 And Row2 <> -1 Or (Row = -1 And Row2 = -1) Then myreplaceGrd2 Row2
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0), , 1)
    If nFound <> -1 Then
         MsgBox "»Ê·Ì’… « ·‘Õ‰ „ÊÃÊœ… ›Ï «·”ÿ— " & nFound
         Exit Sub
    End If
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 3)
    grid1.TextMatrix(grid1.Row, 3) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 4)
    grid1.TextMatrix(grid1.Row, 4) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 6)
    grid1.TextMatrix(grid1.Row, 5) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 7)
    grid1.TextMatrix(grid1.Row, 6) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 8)
    grid1.TextMatrix(grid1.Row, 7) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 9)
    grid1.TextMatrix(grid1.Row, 8) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 10)
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearch_Travel
        CellPos 13, grid1.Row, 2
    Else
        grid1.Select grid1.Rows - 1, 2
    End If
ElseIf ActiveControl.Name = grid2.Name Then
    grid2.TextMatrix(grid2.Row, grid2.Col) = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
    If grid2.Col = 0 Then grid2.TextMatrix(grid2.Row, 1) = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 1)
    Unload oSearchBox
    Grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    Unload oSearchDoc
ElseIf ActiveControl.Name = XBOX.Name Then
    XBOX.Text = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
    xBox_desca.Caption = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 1)
    Unload oSearchBox
    SendKeys "{TAB}"
ElseIf ActiveControl.Name = xFollower.Name Then
    xFollower.BoundText = oSearchDriver.grid1.TextMatrix(oSearchDriver.grid1.Row, 0)
    Unload oSearchDriver
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdCash_Click()
charge_Cashfrm.Show 1
End Sub
Private Sub cmdCost_Click()
cost_fixfrm.Show 1
End Sub
Private Sub cmdDaySales_Click()
casher_closefrm.Show 1
End Sub
Private Sub cmdCargo_Click()
Dim oFlagfrm As New flag_mainfrm, SCODE As String
SCODE = xCargo.BoundText
oFlagfrm.sTable = "CARGO_CODES"
oFlagfrm.sCaption = "«‰Ê«⁄ «·Õ„Ê·« "
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data3.Refresh
xCargo.BoundText = SCODE
If Not xCargo.MatchedWithList Then xCargo.BoundText = ""
End Sub
Private Sub cmdAddTravel_Click()
If Trim(XBOX.Text) = "" Then Exit Sub
Set oAdd.myForm = Me
oAdd.Sbox = XBOX.Text
oAdd.Show 1
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From TRUST_DOC where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRUST_CASH where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRUST_H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
       CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
       If CardTable.BOF Then CardTable.MoveFirst
       myload
    End If
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub

Private Sub cmdGroup_Click()

End Sub

Private Sub CmdInform_Click()
If Option1(1).Value Or Option1(2).Value Then
    Trust_LookupAll Me, oSearchDoc, cFilter & turn(cFilter, " AND ") & IIf(Option1(1).Value, "DONE = 0", "DONE = 1")
Else
    Trust_LookupAll Me, oSearchDoc, cFilter
End If
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
bAddnew = True
mydefine
On Error Resume Next
Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
mysave
On Error Resume Next
Err.Clear
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Command2_Click()
TDaySal.Show 1
End Sub
Private Sub Command3_Click()
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset
loctable.Open "Select * FROM TRAVEL_H", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If Not IsNull(loctable!Time) Then
        dDate = Format(IIf(Val(Format(loctable!Time, "hh")) > 4, loctable!Time, DateAdd("d", -1, loctable!Time)), "YYYY-MM-DD")
        cString = "update TRAVEL_H set TRAVEL_H.date = " & DateSq(dDate)
        cString = cString & turn(cString) & " doc_no = " & MyParn(loctable!doc_no)
        con.Execute cString
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing


Set loctable = New ADODB.Recordset
loctable.Open "Select DOC_NO,SUM(PRICE * QUANT) AS TOTAL FROM TRAVEL GROUP BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    con.Execute "UPDATE TRAVEL_H SET TRAVEL_H.CASH = " & Val(loctable!TOTAL & "") & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop
MsgBox "done..."
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()

End Sub
Private Sub Form_Activate()
On Error Resume Next
If xDoc_No.Tag = LoadMode Then grid1.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1_Validate False
    cmdSave_Click
    KeyCode = 0
ElseIf KeyCode = 115 Then
    itemsgrdfrm.Show 1
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
bEdit = True
openCon con
cList = StrList2("Select code,desca from FILE0_50 WHERE CODE >= '500001'order by desca")

Set grid1.DataSource = data11
data11.ConnectionString = strCon

Set grid2.DataSource = data12
data12.ConnectionString = strCon

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing

closeCon con

Unload oSearch
Unload oSearchDoc
Unload oSearchBox

Set trustfrm = Nothing
Err.Clear
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Trim(XBOX.Text) = "" Then Exit Sub
Travel_Trust_LookupAll Me, oSearch_Travel, XBOX.Text
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEdit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM TRUST_DOC WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        myRemove grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
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
Calctotals
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then MyAddItem

If myreplace(Row) Then
    HandleCntEdit
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
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 2) = ""
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Sub
If Not IsEmpty(aRet) Then
    
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 And Col <> 6 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
If xDone.Value = 1 Or bEdit = False Then
    grid1.Editable = flexEDNone
    Exit Sub
ElseIf .Col = 1 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
'If grid1.Rows < 2 Then Exit Sub
'If grid1.Row = 0 Then
'    grid1.Row = 1
'    grid1.Col = 1
'End If
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    myRemove grid1.Row
End If
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "»Ê·Ì’… «·‘Õ‰|" & "«· «—ÌŒ|" & "„‰|" & "≈·Ì|" & "«·⁄Âœ…|" & "«·„’—Ê›|" & "«·»«ﬁÌ|"
.ColWidth(0) = 500
.ColWidth(1) = 1000
.ColWidth(2) = 1200
.ColWidth(3) = 1400
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1000
.ColWidth(7) = 1050
.ColWidth(8) = 1050
.ColWidth(9) = 1050
'.ColHidden(1) = True
.ColComboList(1) = "..."
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 8 Then
    .Col = Col + 1
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, 1
    .ShowCell Row + 1, 1
Else
    .Select Row, Col
End If
End With
End Sub
Private Sub MyAddItem()
With grid1
.AddItem ""
MakeSerial
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "TRAVEL", addstring(grid1.TextMatrix(I, 1)))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "TRUST_DOC")
        Else
            con.Execute addUpdate(aInsert, "TRUST_DOC", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
cString = "SELECT TRUST_DOC.TRAVEL,TRAVEL_H.POLICY,CONVERT(VARCHAR(10),TRAVEL_H.DATE_POLICY,111),TRAVEL_H.PLACE1,TRAVEL_H.PLACE2,TRAVEL_BAL.TRUST,TRAVEL_BAL.CHARGE,TRAVEL_BAL.BALANCE,TRUST_DOC.ID " & _
          " FROM TRUST_DOC INNER JOIN TRAVEL_H ON TRUST_DOC.TRAVEL = TRAVEL_H.DOC_NO LEFT JOIN TRAVEL_BAL ON (TRUST_DOC.TRAVEL = TRAVEL_BAL.DOC_NO AND TRAVEL_BAL.BOX = " & MyParn(XBOX.Text) & ")"
cString = cString & turn(cString) & " TRUST_DOC.DOC_NO = " & MyParn(xDoc_No.Text)
data11.RecordSource = cString
data11.Refresh
MyAddItem
End With
Calctotals
Fixgrd
End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
openCardTable
myUndo
End Sub
Private Sub xBalance_Change()
'xDone.Enabled = (xDoc_No.Tag = LoadMode & "" And Val(xBalance.Caption) = 0)
'If Val(xBalance.Caption) <> 0 And xDone.Value <> 0 Then
'    MsgBox "·« Ì„ﬂ‰ «· ”ÊÌ… ÊÂ‰«ﬂ —’Ìœ"
'    xDone.Value = 1
'    Exit Function
'End If
End Sub

Private Sub xBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then BoxLookupAll Me, oSearchBox
End Sub
Private Sub xCode_LostFocus()
myLostFocus XCODE
End Sub
Private Sub xCode_sup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub
Private Sub xbox_Validate(Cancel As Boolean)
xBox_desca.Caption = ""
If XBOX.Text = "" Then Exit Sub
XBOX.Text = RetZero(XBOX.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file0_50 where code = " & MyParn(XBOX.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·Œ“‰… €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xBox_desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
Dim I As Integer
If Not bEdit Then Exit Function
If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Trim(XBOX.Text) = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
XBOX.Text = CardTable!BOX & ""
xBox_desca.Caption = CardTable!Box_Desca & ""
xDone.Value = IIf(CardTable!DONE, 1, 0)
bEdit = Not CardTable!DONE
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
myloadgrd2
CellPos2 13, grid2.Rows - 2, grid2.Cols - 1
Handlecontrols LoadMode
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Sub mydefine()
xDoc_No.Text = ""
xTrust_No.Text = ""
xDate.Text = ""
XBOX.Text = ""
xBox_desca.Caption = ""
xClosed.Value = 0
xDone.Value = 0
xTotal_cost.Caption = ""
xTotal_trust.Caption = ""
xRest.Caption = ""
xTotal_P.Caption = ""
xTotal_M.Caption = ""
xBalance.Caption = ""
grid1.Rows = 1
MyAddItem
Fixgrd

grid2.Rows = 1
MyAddItem2
fixgrd2
Handlecontrols DefineMode
HandleState
End Sub
Private Sub Handlecontrols(nMode)
cmdAdd.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDel.Enabled = nMode = LoadMode And bEdit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
XBOX.Enabled = Not (grid1.Rows > 2 Or grid2.Rows > 2)
xDone.Enabled = (xDoc_No.Tag = LoadMode)
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub

Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xDoc_No.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Function Calctotals(Optional nMode As Integer = 0)
Dim nTotal_Cost As Double, nTotal_Trust As Single, nRest As Double
With grid1
For I = 1 To .Rows - 2
    nTotal_Trust = Val(.TextMatrix(I, 6)) + nTotal_Trust
    nTotal_Cost = Val(.TextMatrix(I, 7)) + nTotal_Cost
    nRest = Val(.TextMatrix(I, 8)) + nRest
Next
xTotal_trust.Caption = Myvalue(nTotal_Trust)
xTotal_cost.Caption = Myvalue(nTotal_Cost)
xRest.Caption = Myvalue(nRest)
End With

Dim nTotal_P As Double, nTotal_M As Double, nBalance As Double
With grid2
For I = 1 To .Rows - 2
    nTotal_P = Val(.TextMatrix(I, 2)) + nTotal_P
    nTotal_M = Val(.TextMatrix(I, 3)) + nTotal_M
    nBalance = nBalance + (Val(.TextMatrix(I, 2)) - Val(.TextMatrix(I, 3)))
Next
nBalance = nRest - nBalance
xTotal_P.Caption = Myvalue(nTotal_P)
xTotal_M.Caption = Myvalue(nTotal_M)
xBalance.Caption = Myvalue(nBalance)
HandleState
End With
End Function
Private Function Calctotals2()
'If Val(xDistance.Text) <> 0 Then
'    xGasPerKilo.Caption = Myvalue(Val(xGas.Text) / Val(xDistance.Text))
'Else
'    xGasPerKilo.Caption = ""
'End If
'xGasPerKilo_Differ.Caption = Myvalue((Val(xGasPerKilo.Caption) - Val(xGasPerKilocar.Caption)) * 100)
'xGascar.Caption = Myvalue(Val(xDistance.Text) * Val(xGasPerKilocar.Caption))
'xGas_differ.Caption = Myvalue(Val(xGas.Text) - Val(xGascar.Caption))
End Function

Private Sub CardLookup(Optional pWhere As String = "")
End Sub
Private Function mysave() As Boolean
If Not MYVALID Then Exit Function
Calctotals
If Not myreplace Then Exit Function
Inform " „ Õ›Ÿ «·„” ‰œ"
openCardTable
myUndo
End Function
Private Function doprint() As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1
For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!Date1 = DateFix(xDate.Text)
    temptable!str5 = TurnValue(xtime.Caption)
    
    If xNotes.Text <> "" Then temptable!str2 = "«·”«œ… : " & xNotes.Text
    temptable!Str11 = "iPlanet "
    
    
    temptable!str3 = ArbString(Val(xDoc_No.Text))
    temptable!str6 = .TextMatrix(I, 2)
    temptable!str4 = TurnValue(XBOX.Text)
    temptable!val1 = Val(.TextMatrix(I, 3))
    temptable!val2 = Val(.TextMatrix(I, 4))
    temptable!val3 = Val(.TextMatrix(I, 5))
    temptable!str14 = TurnValue(cComp_Name)
    temptable!str15 = TurnValue(cComp_address)
    temptable!str16 = TurnValue(turn(cComp_Phone, "Phone : ") & cComp_Phone)

    temptable!val4 = Val(xTotal_Item.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xcash.Caption)
    temptable!Val7 = Val(xPay.Caption)
    temptable!Val8 = Val(xRest.Caption)
    temptable.Update
Next I
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.Destination = crptToPrinter
main.REPORT1.ReportFileName = App.Path & "\Reports\sales_bon.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
main.REPORT1.Destination = crptToWindow
doprint = True
GoTo closeCon
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon:
temptable.Close
Set temptable = Nothing
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
cmdSave.Enabled = (bEdit)
XBOX.Enabled = Not (grid1.Rows > 2 Or grid2.Rows > 2)
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TRUST_H.*,FILE0_50.DESCA AS BOX_DESCA FROM TRUST_H INNER JOIN FILE0_50 ON TRUST_H.BOX = FILE0_50.CODE"
If Option1(1).Value Then
    cString = cString & turn(cString) & "TRUST_H.DONE = 0"
ElseIf Option1(2).Value Then
    cString = cString & turn(cString) & "TRUST_H.DONE = 1"
End If
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myerror
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
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With grid2
    If KeyCode = 46 And .Row <> .Rows - 1 And bEdit Then
        If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.BeginTrans
                On Error GoTo myerror
                con.Execute "DELETE FROM Trust_Cash WHERE ID = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            myRemove2 .Row
        End If
    ElseIf KeyCode = 13 Then
        CellPos2 KeyCode, .Row, .Col
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd2
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid2
Calctotals
If Row = .Rows - 2 Then
    If .TextMatrix(.Rows - 2, 0) <> "" Then
        .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
        .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
    End If
End If
If Not validRow2(Row) Then Exit Sub
If Row = .Rows - 1 Then MyAddItem2

If myreplace(, Row) Then
    HandleCntEdit
    If .TextMatrix(Row, .Cols - 1) = "" Then
        myloadgrd2
    End If
Else
    myloadgrd2
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd2
End Sub
Private Sub grid2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 Then
    CellPos2 KeyCode, Row, Col
End If
End Sub
Private Function validRow2(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 1)) Then Exit Function
If (Not IsNumeric(.TextMatrix(Row, 2))) And Not IsNumeric(.TextMatrix(Row, 3)) Then Exit Function
End With
validRow2 = True
End Function
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid2.Rows - 1 And OldRow <> 0 And grid2.TextMatrix(OldRow, grid2.Cols - 1) = "" Then
    If Not validRow2(OldRow) Then
        myRemove2 OldRow
    End If
End If
End Sub
Private Sub grid2_EnterCell()
With grid2
If xClosed.Value = 1 Or bEdit = False Then
    .Editable = flexEDNone
    Exit Sub
End If
If (.Col = 0 Or .Col = 1 Or .Col = 2 Or .Col = 3) Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub GRID2_GotFocus()
If grid2.Rows < 2 Then Exit Sub
If grid2.Row = 0 Then
    grid2.Row = 1
    grid2.Col = 1
End If
grid2_EnterCell
End Sub
Private Sub grid2_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow2(grid2.Row) And grid2.Row <> grid2.Rows - 1 And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "" Then
    myRemove2 grid2.Row
End If
End Sub
Private Sub GRID2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid2
If Col = 0 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 1 Then
    .EditText = Format(.EditText, "YYYY/MM/DD")
    If Not IsDate(.EditText) Then
        MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Val(.EditText) <> 0 And Val(.TextMatrix(Row, 3)) <> 0 Then
        .TextMatrix(Row, 3) = ""
    End If
ElseIf Col = 3 Then
    If Val(.EditText) <> 0 And Val(.TextMatrix(Row, 2)) <> 0 Then
        .TextMatrix(Row, 2) = ""
    End If
End If
End With
End Sub
Private Sub fixgrd2()
With grid2
.FormatString = "«·Œ“‰…|" & "«· «—ÌŒ|" & "≈Ìœ«⁄|" & "”Õ»|"
.ColWidth(0) = 3000
.ColWidth(1) = 1500
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColComboList(0) = cList
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid2
KeyCode = 0
If Col < .Cols - 2 Then
    .Col = Col + 1
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid2, Row + 1, 0, 2)
    .ShowCell Row + 1, 1
Else
    .Select Row, Col
End If
End With
End Sub
Private Sub MyAddItem2()
With grid2
.AddItem ""
If .Rows > 2 Then
    If .TextMatrix(.Rows - 2, 0) <> "" Then .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
End If
End With
End Sub
Private Function myreplaceGrd2(Row) As Boolean
Dim aInsert As Variant
With grid2
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[BOX]", addstring(.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "[VALUE_P]", Val(.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "[VALUE_M]", Val(.TextMatrix(I, 3)))
        If .TextMatrix(I, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "TRUST_CASH")
        Else
            con.Execute addUpdate(aInsert, "TRUST_CASH", "ID = " & .TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd2 = True
End Function
Private Sub myloadgrd2()
With grid2
cString = "SELECT TRUST_CASH.BOX,CONVERT(VARCHAR(10),TRUST_CASH.DATE,111),TRUST_CASH.VALUE_P,TRUST_CASH.VALUE_M,ID " & _
          " FROM TRUST_CASH"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
data12.RecordSource = cString
data12.Refresh
MyAddItem2
End With
Calctotals
fixgrd2
End Sub
Private Sub xtotal_GotFocus()
myGotFocus xTotal
End Sub
Private Sub xtotal_LostFocus()
myLostFocus xTotal
Calctotals
End Sub
Private Sub xGas_GotFocus()
myGotFocus xGas
End Sub
Private Sub xGas_LostFocus()
myLostFocus xGas
Calctotals2
End Sub
Private Sub xCar_GotFocus()
myGotFocus xCar
End Sub
Private Sub xCar_LostFocus()
myLostFocus xCar

End Sub
Private Sub xDistance_GotFocus()
myGotFocus xDistance
End Sub
Private Sub xDistance_LostFocus()
myLostFocus xDistance
Calctotals2
End Sub
Private Sub xPlace2_GotFocus()
myGotFocus xPlace2
End Sub
Private Sub xPlace2_LostFocus()
myLostFocus xPlace2
End Sub
Private Sub xPlace1_GotFocus()
myGotFocus xPlace1
End Sub
Private Sub xPlace1_LostFocus()
myLostFocus xPlace1
End Sub
Private Sub xCode_GotFocus()
myGotFocus XCODE
End Sub
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
Private Sub xDriver_GotFocus()
myGotFocus xDriver
End Sub
Private Sub xDriver2_GotFocus()
myGotFocus xDriver2
End Sub
Private Sub xDriver2_LostFocus()
myLostFocus xDriver2
End Sub
Private Sub xWeight_GotFocus()
myGotFocus xWeight
End Sub
Private Sub xWeight_LostFocus()
myLostFocus xWeight
Calctotals
End Sub
Private Sub xFollower_GotFocus()
myGotFocus xFollower
End Sub
Private Sub xFollower_LostFocus()
myLostFocus xFollower
If Not xFollower.MatchedWithList Then xFollower.BoundText = ""
End Sub
Private Function Retlist() As String
Dim alist As Variant
If xDriver.MatchedWithList Then
    alist = AddFlag(alist, "CODE", xDriver.BoundText)
    alist = AddFlag(alist, "DESCA", xDriver.Text)
    aString = AddFlag(aString, alist)
End If
If xDriver2.MatchedWithList Then
    alist = AddFlag(Empty, "CODE", xDriver2.BoundText)
    alist = AddFlag(alist, "DESCA", xDriver2.Text)
    aString = AddFlag(aString, alist)
End If
Retlist = StrListArray(aString)
If Retlist = "" Then Retlist = cList
End Function
Private Sub xCargo_GotFocus()
myGotFocus xCargo
End Sub
Private Sub xCargo_LostFocus()
myLostFocus xCargo
If Not xCargo.MatchedWithList Then xCargo.BoundText = ""
End Sub
Private Sub xTotal_sup_GotFocus()
myGotFocus xTotal_sup
End Sub
Private Sub xTotal_sup_LostFocus()
myLostFocus xTotal_sup
Calctotals
End Sub
Private Sub xCode_sup_GotFocus()
myGotFocus xCode_sup
End Sub
Private Sub xCode_sup_LostFocus()
myLostFocus xCode_sup
Calctotals
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial
Handlecontrols xDoc_No.Tag
Calctotals
End Sub
Private Sub myRemove2(Row As Long)
grid2.RemoveItem Row
Handlecontrols xDoc_No.Tag
Calctotals
End Sub
Sub Addproc()
With oAdd.grid1
For I = 1 To .Rows - 1
    If Val(.TextMatrix(I, .Cols - 1)) <> 0 Then
        grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
        grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(I, 0)
        grid1.TextMatrix(grid1.Rows - 1, 2) = .TextMatrix(I, 3)
        grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(I, 4)
        grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(I, 5)
        grid1.TextMatrix(grid1.Rows - 1, 5) = .TextMatrix(I, 6)
        grid1.TextMatrix(grid1.Rows - 1, 6) = .TextMatrix(I, 7)
        grid1.TextMatrix(grid1.Rows - 1, 7) = .TextMatrix(I, 8)
        grid1.TextMatrix(grid1.Rows - 1, 8) = .TextMatrix(I, 9)
        grid1.AddItem ""
    End If
Next
End With
Unload oAdd
mysave
'On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub
Private Sub HandleState()
If Val(xTotal_trust.Caption) <> 0 Or Val(xTotal_cost.Caption) <> 0 Then
    lblState.Visible = True
    If xDone.Value = 1 Then
        lblState.ForeColor = &H8000&
        lblState.Caption = " „   ”ÊÌ… «·⁄„·Ì…"
    Else
        lblState.ForeColor = vbRed
        lblState.Caption = "·„ Ì „  ”ÊÌ… «·⁄„·Ì…"
    End If
Else
    lblState.Visible = False
End If
End Sub
Private Sub xDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
con.Execute "UPDATE TRUST_H SET DONE = " & xDone.Value & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
Inform " „ «· ⁄œÌ· »‰Ã«Õ"
openCardTable
myUndo
End Sub

