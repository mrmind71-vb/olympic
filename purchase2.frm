VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form purchasefrm2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15285
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
   ScaleHeight     =   9360
   ScaleWidth      =   15285
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1140
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   630
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Picture         =   "purchase2.frx":2579
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   5595
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   90
         Picture         =   "purchase2.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   180
         Width           =   1410
      End
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1500
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":6D48
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2820
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":95E2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4140
         Picture         =   "purchase2.frx":BB8E
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   765
      Width           =   10590
      Begin VB.CommandButton cmdClient 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   540
         Width           =   330
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8280
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
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
         Height          =   330
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   345
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   345
         Left            =   6660
         TabIndex        =   34
         Top             =   540
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Index           =   1
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   585
         Width           =   615
      End
      Begin VB.Label xBalance 
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
         Height          =   375
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9465
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
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
         Index           =   0
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   675
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3465
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -1980
      Top             =   990
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
      Left            =   -720
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
      Left            =   -720
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6000
      Left            =   90
      TabIndex        =   3
      Top             =   1800
      Width           =   15090
      _cx             =   26617
      _cy             =   10583
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
   Begin VB.Frame fmTotal 
      Height          =   1320
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7830
      Width           =   11175
      Begin VB.TextBox xRateDis 
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
         Height          =   330
         Left            =   7695
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   540
         Width           =   645
      End
      Begin VB.TextBox xRateTax 
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
         Height          =   330
         Left            =   4050
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox xDiscount 
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
         Height          =   330
         Left            =   8370
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox xTax 
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
         Height          =   330
         Left            =   4905
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   195
      End
      Begin VB.Label xTotalFix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   330
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "«·≈Ã„«·Ì :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·Œ’„ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì «·«’‰«› :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   330
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "«·÷—Ì»… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblTotalQuant 
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label10 
         Caption         =   "«·«Ã„«·Ì :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   330
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   1860
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   7830
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   37
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
         Picture         =   "purchase2.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase2.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   38
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
         Picture         =   "purchase2.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase2.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   39
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
         Picture         =   "purchase2.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase2.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   40
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
         Picture         =   "purchase2.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase2.frx":1CE87
      End
   End
End
Attribute VB_Name = "purchasefrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Integer
Dim con As New ADODB.Connection
Public bEdit As Boolean, nRound As Integer, nRoundG As Integer
Dim CardTable As ADODB.Recordset, cFile As String, cFileHeader As String
Dim oSearchDoc As New Search3, osearchitem As New Search3, oSearchSup As New Search3
Public sDoc_no As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(7, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(XCODE.BoundText)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.Text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "Discount"
aInsert(4, 1) = Val(xDiscount.Text)

aInsert(5, 0) = "Tax"
aInsert(5, 1) = Val(xTax.Text)

aInsert(6, 0) = "rateDis"
aInsert(6, 1) = Val(xRateDis.Text)

aInsert(7, 0) = "RateTax"
aInsert(7, 1) = Val(xRateTax.Text)

con.BeginTrans
On Error GoTo myerror
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 1 Then
        Dim bNew As Boolean: bNew = grid1.Row = grid1.Rows - 1
        grid1.TextMatrix(grid1.Row, 1) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 1)
        grid1.TextMatrix(grid1.Row, 3) = 1
        Grid1_AfterEdit grid1.Row, grid1.Col
        If bNew Then grid1.Row = grid1.Rows - 1
        If grid1.Row < grid1.Rows - 1 Then osearchitem.Hide
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    oSearchDoc.Hide
ElseIf ActiveControl.Name = XCODE.Name Then
    ActiveControl.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
End Sub

Private Sub CmdAdditem_Click()

End Sub
Private Sub cmdClient_Click()
Dim nCode As String
nCode = XCODE.BoundText
Clients.myFlag = 2
Clients.Show 1
DATA2.Refresh
XCODE.BoundText = nCode
If Not XCODE.MatchedWithList Then XCODE.BoundText = ""
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    ' Õ–› «·„” ‰œ
    con.Execute "Delete From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
       CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
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
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub

Private Sub xCode_Validate(Cancel As Boolean)
With XCODE
    If Not .MatchedWithList Then .BoundText = ""
End With
End Sub

Private Sub xTax_LostFocus()
Calctotals
End Sub
Private Sub xRateDis_LostFocus()
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRateDis.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
        xDiscount.Text = Round((Val(xRateDis.Text) * Val(xTotalItem.Caption)) / 100, nRound)
    End If
Else
    xDiscount.Text = ""
End If
Calctotals
End Sub
Private Sub xRateTax_LostFocus()
If Val(xTotalFix.Caption) <> 0 Then
    If Round(Val(xRateTax.Text), nRound) <> Round(Val(xTax.Text) / Val(xTotalFix.Caption) * 100, nRound) Or xTax.Locked Then
        xTax.Text = Round((Val(xRateTax.Text) * Val(xTotalFix.Caption)) / 100, nRound)
    End If
Else
   xTax.Text = ""
End If
Calctotals
End Sub
Private Sub cmdExit_Click()
Unload Me
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
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
mysave
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
nRound = 2
nRoundG = 3
openCon con
bEdit = True
Select Case myPublic
Case 0
    cFile = "File7_20"
    cFileHeader = "File7_20H"
    Me.Caption = "›« Ê—… „‘ —Ì« "
    cFileSerial = "File7_22"
Case 1
    cFile = "FILE7_30"
    cFileHeader = "File7_30H"
    Me.Caption = "›« Ê—… „—œÊœ „‘ —Ì« "
End Select

Dim cString As String
cString = "SELECT " & cFileHeader & ".*,FILE4_10.DESCA AS CODEDESCA FROM " & cFileHeader & " INNER JOIN FILE4_10 ON " & cFileHeader & ".CODE = FILE4_10.CODE"
cString = cString & " Order by " & cFileHeader & ".DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE4_10"
Set XCODE.RowSource = DATA2
XCODE.ListField = "Desca"
XCODE.BoundColumn = "Code"

Set grid1.DataSource = data3
data3.ConnectionString = strCon
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If grid1.Col = 1 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    MakeSerial Row
End If
Calctotals
End With
End Sub
Private Sub Grid1_EnterCell()
'If (grid1.Col = 2) Then
'    grid1.Editable = flexEDNone
'Else
    grid1.Editable = flexEDKbdMouse
'End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.Select 1, 1
End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub
Private Function MYVALID() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If XCODE.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
For I = 1 To grid1.Rows - 2
    If Not validRow(I) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xStore.BoundText = CardTable!store & ""
XCODE.BoundText = CardTable!Code & ""
xDiscount.Text = Myvalue(Val(CardTable!Discount & ""), "00.00")
xTax.Text = Myvalue(Val(CardTable!tax & ""), "00.00")
xRateTax.Text = CardTable!Ratetax & ""
xRateDis.Text = CardTable!RateDis & ""
myloadgrd
Handlecontrols LoadMode
grid1.ShowCell grid1.Rows - 1, 1
grid1.Select grid1.Rows - 1, 1
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
xDate.Text = Format(Date, "YYYY-MM-DD")
xStore.BoundText = sDefStore
xStore.Enabled = xStore.BoundText = ""
xBalance.Caption = ""
XCODE.BoundText = ""
xDiscount.Text = ""
xTotal.Caption = ""
xTax.Text = ""
xTotalItem.Caption = ""
xDiscount.Text = ""
xRateDis.Text = ""
xTax.Text = ""
xRateTax.Text = ""
xtotalQuant.Caption = ""
grid1.Rows = 1
grid1.AddItem ""
Fixgrd
MakeSerial 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bEdit)
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDiscount_LostFocus()
Calctotals
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRow(grid1.Row) Then
    grid1.AddItem "", grid1.Row
    MakeSerial grid1.Row - 1
ElseIf KeyCode = 112 Then
    ItemsLookupAll Me, osearchitem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEdit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        Calctotals
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 2) = ""
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
Dim aRet As Variant
aRet = aGetDesca("select desca from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 1)))
If UBound(aRet) <> 0 Then
    grid1.TextMatrix(Row, 2) = aRet(1)
    grid1.TextMatrix(Row, 4) = LastPrice(grid1.TextMatrix(Row, 1))
End If
End Sub
Private Function Calctotals()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalDis As Double
With grid1
For I = 1 To grid1.Rows - 2
    .TextMatrix(I, 5) = Myvalue(Round(Val(.TextMatrix(I, 3)) / Val(.TextMatrix(I, 4)), nRound))
    nTotalitem = nTotalitem + Val(.TextMatrix(I, 5))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 3))
Next

If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRateDis.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xRateDis.Text = Round((Val(xDiscount.Text) / Val(xTotalItem.Caption)) * 100, nRound)
    End If
Else
    xRateDis.Text = ""
End If

xTotalItem.Caption = nTotalitem
xTotalFix.Caption = nTotalitem - Val(xDiscount.Text)


If Val(xTotalFix.Caption) <> 0 Then
    If Round(Val(xRateTax.Text), nRound) <> Round(Val(xTax.Text) / Val(xTotalFix.Caption) * 100, nRound) Or xTax.Locked Then
        xTax.Text = Round((Val(xRateTax.Text) * Val(xTotalFix.Caption)) / 100, nRound)
    End If
Else
    xTax.Text = ""
End If
xTotal.Caption = Val(xTotalItem.Caption) + Val(xTax.Text) - Val(xDiscount.Text)
xtotalQuant.Caption = Round(nTotalQuant, 2)
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO, CONVERT(VARCHAR(10), " & cFileHeader & ".[DATE],111), FILE4_10.Desca " & _
                  " FROM  (" & cFileHeader & " INNER JOIN FILE4_10 ON " & cFileHeader & ".CODE " & " = FILE4_10.CODE )"
                  
Generalarray(2) = "Order by " & cFileHeader & ".Date,DOC_NO "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6  or  %%FILE4_10.DESCA%% OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "≈”„ " & cCodeDesca
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
If Trim(grid1.TextMatrix(nRow, 2)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub MakeSerial(Optional nBeginRow As Long = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "«·’‰›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ì|" & "„·«ÕŸ« |"
.ColWidth(0) = 600
.ColWidth(1) = 2000
.ColWidth(2) = 5500
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 3000
.ColHidden(.Cols - 1) = True
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From " & cFileClient
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub
Private Sub doprint(nPrint)
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    temptable.AddNew
    If nPrint = 1 Then
        If myPublic = 0 Then
            temptable!str21 = "„” ‰œ „‘ —Ì«   "
        Else
            temptable!str21 = "„” ‰œ „—œÊœ „‘ —Ì«   "
        End If
    Else
        If myPublic = 0 Then
            temptable!str21 = "„” ‰œ ≈–‰ œŒÊ· »÷«⁄…  "
        Else
            temptable!str21 = "„” ‰œ ≈–‰ Œ—ÊÃ »÷«⁄…"
        End If
    End If
    temptable!str1 = xDoc_No.Text
    temptable!str2 = xDate.Text
    temptable!str3 = Format(XCODE.BoundText)
    temptable!str4 = xCodeDesca.Caption
    temptable!str5 = xStore.Text
    temptable!Str11 = TurnValue(grid1.TextMatrix(I, 1))
    
    temptable!str12 = TurnValue(grid1.TextMatrix(I, 2))
    temptable!val1 = Val(grid1.TextMatrix(I, 3))
    temptable!val10 = I
    temptable!val11 = IIf(Not isOut, 0, 1)
    
       
    temptable!val12 = Val(GetDesca("Select Package from file1_10 where item = " & MyParn(grid1.TextMatrix(I, 1))))
    temptable!val5 = Val(xTotalItem.Caption)
    temptable!Val6 = Val(xDisItem.Caption)
    temptable!Val7 = Val(xDiscount.Text)
    temptable!Val8 = Val(xTotal.Caption)
    temptable!val9 = Val(xtotalQuant.Caption)
    
    temptable!str10 = MyOnly(Val(xTotal.Caption))
    temptable!val9 = myPublic
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If nPrint = 1 Then
    mainfrm.Report1.ReportFileName = App.Path & "\Reports\purchase.rpt"
Else
    mainfrm.Report1.ReportFileName = App.Path & "\Reports\purchase_A.rpt"
End If
mainfrm.Report1.DataFiles(0) = tempFile
mainfrm.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub mysave(Optional bMsg As Boolean = True)
If Not MYVALID Then Exit Sub
Calctotals
If Not myreplace Then Exit Sub
If bMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Sub myproc2(nDoc_no)
bNoMsgExit = True
xDoc_No.Text = ndoc_no2
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub myreplaceGrd()
Dim aInsert(7, 1)
With grid1
    For I = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(I, 1))
        
       
        aInsert(3, 0) = "quant"
        aInsert(3, 1) = .TextMatrix(I, 5 - 2)

        aInsert(4, 0) = "Price"
        aInsert(4, 1) = Val(.TextMatrix(I, 6 - 2))

        aInsert(5, 0) = "Total"
        aInsert(5, 1) = Val(.TextMatrix(I, 7 - 2))

        aInsert(6, 0) = "notes"
        aInsert(6, 1) = addstring(.TextMatrix(I, 8 - 2))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub xTax_GotFocus()
xTax.SelStart = 0
xTax.SelLength = Len(xTax.Text)
End Sub
Private Sub xDiscount_GotFocus()
xDiscount.SelStart = 0
xDiscount.SelLength = Len(xDiscount.Text)
End Sub
Private Sub xRateDis_GotFocus()
xRateDis.SelStart = 0
xRateDis.SelLength = Len(xRateDis.Text)
End Sub
Private Sub xRateTax_GotFocus()
xRateTax.SelStart = 0
xRateTax.SelLength = Len(xRateTax.Text)
End Sub
Private Sub xCode_GotFocus()
XCODE.SelStart = 0
XCODE.SelLength = Len(XCODE.BoundText)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xusername_GotFocus()
xusername.SelStart = 0
xusername.SelLength = Len(xusername.Text)
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
    Calctotals
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        Calctotals
    End If
End If
End With
End Sub
Private Sub myloadgrd()
With grid1
    If myPublic = 0 Then
        Dim cFieldDesca As String
        cString = "SELECT FILE7_20.ITEM, FILE1_10.DESCA, QUANT,FILE7_20.Price,FILE7_20.Total,FILE7_20.NOTES,FILE7_20.ID" & _
              " FROM FILE7_20 INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM"
        cString = cString & turn(cString) & " FILE7_20.DOC_NO = " & MyParn(xDoc_No.Text) & " order by FILE7_20.ID"
    Else
        cString = "SELECT FILE7_30.ITEM, FILE1_10.DESCA, QUANT,FILE7_30.Price,FILE7_30.Total,FILE7_30.NOTES,FILE7_30.ID" & _
              " FROM FILE7_30 INNER JOIN FILE1_10 ON FILE7_30.ITEM = FILE1_10.ITEM"
        cString = cString & turn(cString) & " FILE7_30.DOC_NO = " & MyParn(xDoc_No.Text) & " order by FILE7_30.ID"
    End If
    data3.RecordSource = cString
    data3.Refresh
    grid1.AddItem ""
    MakeSerial
End With
Calctotals
Fixgrd
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
KeyCode = 0
If Col = 1 Then
    grid1.Col = 3
    If Not grid1.RowIsVisible(Row) Then grid1.ShowCell Row, 1
ElseIf Col = 3 Or Col = 4 Then
    grid1.Col = Col + 1
ElseIf Col = 5 Then
    If Row < grid1.Rows - 1 Then
        grid1.Row = Row + 1
        grid1.Col = 1
        If Not grid1.RowIsVisible(Row) Then grid1.ShowCell Row, 1
    End If
End If
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT " & cFileHeader & ".*,FILE4_10.DESCA AS CODEDESCA FROM " & cFileHeader & " INNER JOIN FILE4_10 ON " & cFileHeader & ".CODE = FILE4_10.CODE"
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
cString = cString & " Order by " & cFileHeader & ".DOC_NO"
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

