VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salaryfrm 
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
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   3375
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1125
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   675
      Width           =   1275
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salary.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "salary.frx":2579
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
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
      TabIndex        =   22
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salary.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salary.frx":6CFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salary.frx":9594
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "salary.frx":BB40
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   1545
      Begin VB.CommandButton CMDPRINT 
         Height          =   510
         Left            =   45
         Picture         =   "salary.frx":E313
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
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
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   8925
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
         Left            =   5805
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
         Left            =   7155
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   780
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
         TabIndex        =   2
         Top             =   180
         Width           =   1725
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
         TabIndex        =   3
         Top             =   585
         Width           =   7845
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
         Left            =   2610
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   960
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label7 
         Caption         =   "«·»Ì«‰ :"
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
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   675
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
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   660
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   900
      Top             =   1215
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
      Left            =   2835
      Top             =   315
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1125
      Top             =   1080
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin VB.Frame fmTotal 
      Height          =   960
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   8460
      Width           =   6315
      Begin VB.TextBox xTax 
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
         Left            =   3060
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label2 
         Caption         =   "’«›Ì «·„— »"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   540
         Width           =   1560
      End
      Begin VB.Label Label10 
         Caption         =   "«” ﬁÿ«⁄« "
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1650
      End
      Begin VB.Label Label5 
         Caption         =   "«÷«›« "
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   690
      End
      Begin VB.Label xTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1560
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6630
      Left            =   135
      TabIndex        =   30
      Top             =   1800
      Width           =   15765
      _cx             =   27808
      _cy             =   11695
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
      Cols            =   6
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
      TabIndex        =   32
      Top             =   8415
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   33
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
         Picture         =   "salary.frx":1073D
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salary.frx":1290D
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   34
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
         Picture         =   "salary.frx":14A55
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salary.frx":16C1D
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   35
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
         Picture         =   "salary.frx":18D6C
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salary.frx":1AF4C
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   36
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
         Picture         =   "salary.frx":1D0A7
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salary.frx":1F263
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8415
      Visible         =   0   'False
      Width           =   3165
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   15
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
Attribute VB_Name = "salaryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bEdit As Boolean
Dim formMode, dDateLast As String
Dim oSearch As New Search3, oSearchEmp As New Search3
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[date]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[YEAR]", Val(xYear.Text))
aInsert = AddFlag(aInsert, "[MONTH]", Val(xMonth.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Text = "" Then
    xDoc_No.Text = Newflag("FILE2_50H", "doc_no")
    aInsert = AddFlag(aInsert, "[DOC_NO]", addvalue(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE2_50H")
Else
    con.Execute addUpdate(aInsert, "FILE2_50H", "doc_no = " & xDoc_No.Text)
End If
myreplaceGrd Row
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
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearchEmp
        CellPos 13, grid1.Row, 2
    Else
        grid1.Select grid1.Rows - 1, 1
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload Search
    myUndo
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
    con.Execute "Delete  From FILE2_50 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE2_50H where Doc_No = " & MyParn(xDoc_No.Text)
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
Private Sub cmdPrint_Click()
Dim cHeader1 As String
cHeader1 = "Õ÷Ê— «·ÌÊ„Ì… ⁄‰ ‘Â— " & xMonthDesca.Caption & Space(1) & "”‰… " & xYear.Text
grid1.Rows = grid1.Rows - 1
PrintGrdNew.doprint grid1, 0.8, -3, cHeader1, , , , False, False, 10
PrintGrdNew.Show 1
MyAddItem
End Sub
Private Sub Command1_Click()
xDoc_No.Text = ""
xDate.Text = Format(Date, "YYYY-MM-DD")
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
If Not (IsNumeric(xMonth.Text) And IsNumeric(xYear.Text) And IsDate(xDate.Text)) Then
    MsgBox "«·‘Â— Ê«·”‰… «Ê«· «—ÌŒ €Ì— „”Ã·Ì‰"
    Exit Sub
End If

If Not (Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12) Then
    MsgBox "«·‘Â— €Ì— ’ÕÌÕ"
    Exit Sub
End If

cString = "select code,desca  from driver where isday = 1"
cString = cString & turn(cString) & "((date_end is null) or Date_end >= " & DateSq(xDate.Text) & ")"
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If grid1.FindRow(loctable!Code & "", , 1) = -1 Then
        grid1.TextMatrix(grid1.Rows - 1, 1) = loctable!Code
        grid1.TextMatrix(grid1.Rows - 1, 2) = loctable!Desca
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
If xDoc_No.Tag = LoadMode Then
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
bEdit = True
openCon con
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
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
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
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.Tag = LoadMode
        xDoc_No.Enabled = False
    End If
    If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadgrd
    End If
End If
End With
End Sub
Private Sub Grid1_EnterCell()
If (grid1.Col = 1 Or grid1.Col = 3 Or grid1.Col = 4) Then
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
If Val(xMonth) < 1 Or Val(xMonth) > 12 Then
    If Not bIgMsg Then MsgBox "«·‘Â— €Ì— ”·Ì„"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xMonth.Text = CardTable!Month & ""
xMonthDesca.Caption = arbMonth(Val(CardTable!Month & ""))
xYear.Text = CardTable!Year & ""
xDesca.Text = CardTable!Desca & ""
myloadgrd
Handlecontrols LoadMode
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub mydefine()
xDoc_No.Text = ""
xDate.Text = ""
xMonthDesca.Caption = ""
xMonth.Text = ""
xYear.Text = Year(Date)
xDesca.Text = ""
grid1.Rows = 1
MyAddItem
Fixgrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdPrint.Enabled = (nMode = LoadMode)
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And doc_no = ""
xMonth.Enabled = (nMode = DefineMode)
xYear.Enabled = (nMode = DefineMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
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
            aRet = GetFields("select * from driver where isDay = 1 and code = " & MyParn(grid1.EditText), con)
            If IsEmpty(aRet) Then
               MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
               Cancel = True
            Else
                grid1.TextMatrix(Row, 2) = retFlag(aRet, "desca") & ""
            End If
        End If
    End If
End If
End Sub
Private Sub XDATE_DblClick()
Set datefrm.oDate = xDate
datefrm.Show 1
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xDoc_No.Tag = LoadMode Then
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
    DriverLookupAll Me, oSearchEmp, "isday = 1"
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
'    grid1.TextMatrix(i, 8) = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7))
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
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO, [Month] & '-' & [Year],convert(VARCHAR(10),[DATE],111),Desca " & _
                  " FROM  FILE2_50H"

Generalarray(2) = "Order by Year,Month,DOC_NO "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·‘Â—-«·”‰…-«· «—ÌŒ-«·Ê’›"
listarray(0, 1) = "( **[year]** OR  **[Month]** or %%Desca%% " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·‘Â—-«·”‰…"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1200

GrdArray(3, 0) = "»Ì«‰"
GrdArray(3, 1) = 3000

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
.FormatString = "„|" & "ﬂÊœ|" & "«·«”„|" & "«Ì«„ «·Õ÷Ê—|" & "„·«ÕŸ« |"
.ColWidth(0) = 500
.ColWidth(1) = 1000
.ColWidth(2) = 6000
.ColWidth(3) = 1000
.ColWidth(4) = 5000
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
        
        temptable!val3 = Val(grid1.TextMatrix(I, 4))
        
        temptable!val5 = Val(grid1.TextMatrix(I, 5))
        temptable!Val6 = Val(grid1.TextMatrix(I, 6))
        temptable!Val7 = Val(grid1.TextMatrix(I, 7))
        
        temptable!Val8 = Val(grid1.TextMatrix(I, 8))
        temptable!val9 = Val(grid1.TextMatrix(I, 9))
        temptable!val10 = Val(grid1.TextMatrix(I, 10))
        temptable!val11 = Val(grid1.TextMatrix(I, 11))
        temptable!val12 = Val(grid1.TextMatrix(I, 12))
        temptable!str2 = TurnValue(grid1.TextMatrix(I, 13))
         
        temptable!str3 = TurnValue(xDesca_w1.Text)
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
main.Report1.ReportFileName = App.Path & "\Reports\salary2.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function myreplaceGrd(Row As Long) As Boolean
Dim aInsert As Variant
With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "DAYS", Val(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(I, 4)))
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
    cString = "SELECT FILE2_50.CODE,DRIVER.DESCA, FILE2_50.DAYS,FILE2_50.NOTES,FILE2_50.ID" & _
              " FROM FILE2_50 inner JOIN DRIVER ON FILE2_50.CODE = DRIVER.CODE WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    DATA11.RecordSource = cString
    DATA11.Refresh
    grid1.AddItem ""
End With
Calctotals
Fixgrd
MakeSerial
End Sub
Private Sub xMonth_Validate(Cancel As Boolean)
If Trim(xMonth.Text) = "" Then Exit Sub
If IsNumeric(xMonth.Text) Then
    If xMonth.Text < 1 Or xMonth.Text > 12 Then
        MsgBox "«·‘Â— ·« Ì’·Õ"
        Cancel = True
        Exit Sub
    End If
    xMonthDesca.Caption = arbMonth(xMonth.Text)
    If IsNumeric(xMonth.Text) Then
        If Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12 Then
            Dim sDoc_no As Variant
            sDoc_no = GetField("Select doc_no from FILE2_50H where [YEAR] = " & xYear.Text & _
                " and [Month] = " & xMonth.Text)
            If Not IsEmpty(sDoc_no) Then
                xDoc_No.Text = sDoc_no
                myUndo
            ElseIf xDoc_No.Tag = LoadMode Then
                mydefine
            End If
        End If
    End If
End If
End Sub
Private Sub xYear_Validate(Cancel As Boolean)
If Trim(xYear.Text) = "" Then Exit Sub
If IsNumeric(xYear.Text) Then
    If Val(xYear.Text) < 2010 Or Val(xYear.Text) > 2030 Then
        MsgBox "«·”‰… ·«  ’·Õ"
        Cancel = True
        Exit Sub
    End If
    If IsNumeric(xMonth.Text) Then
        If Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12 Then
            Dim sDoc_no As Variant
            sDoc_no = GetField("Select doc_no from FILE2_50H where [YEAR] = " & xYear.Text & _
                " and [Month] = " & xMonth.Text)
            If Not IsEmpty(sDoc_no) Then
                xDoc_No.Text = sDoc_no
                myUndo
            Else
                mydefine
            End If
        End If
    End If
End If
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM FILE2_50H"
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
cString = cString & " Order by FILE2_50H.DOC_NO"
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
If Col < grid1.Cols - 3 Then
    grid1.Col = Col + 1 + IIf(Col = 1, 1, 0)
ElseIf Row < grid1.Rows - 1 Then
    Dim nCol As Long
    nCol = NextEmpty(grid1, Row + 1, 1, 3)
    grid1.Select Row + 1, IIf(nCol > 1, 3, nCol)
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

