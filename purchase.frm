VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form purchasefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8970
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
   ScaleHeight     =   8970
   ScaleWidth      =   15285
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1140
      Left            =   3195
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   945
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "purchase.frx":2579
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   6945
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   1485
         Picture         =   "purchase.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "purchase.frx":6D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   135
         Width           =   1410
      End
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase.frx":9172
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase.frx":BA0C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5580
         Picture         =   "purchase.frx":DFB8
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   720
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
         TabIndex        =   35
         Top             =   585
         Visible         =   0   'False
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
         Height          =   375
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
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   330
         Left            =   6660
         TabIndex        =   2
         Top             =   585
         Width           =   2715
         _ExtentX        =   4789
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   330
         Left            =   6660
         TabIndex        =   4
         Top             =   945
         Width           =   2715
         _ExtentX        =   4789
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
      Begin VB.Label lblBox 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   990
         Width           =   600
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
         TabIndex        =   36
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Visible         =   0   'False
         Width           =   2130
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      Left            =   -405
      Top             =   630
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
      Height          =   5415
      Left            =   90
      TabIndex        =   5
      Top             =   2115
      Width           =   15090
      _cx             =   26617
      _cy             =   9551
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
      TabIndex        =   13
      Top             =   7560
      Width           =   11175
      Begin VB.TextBox xRate 
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
         TabIndex        =   29
         Top             =   540
         Width           =   645
      End
      Begin VB.TextBox xRate1 
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
         TabIndex        =   28
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox xTax1 
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
         TabIndex        =   6
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
         TabIndex        =   34
         Top             =   225
         Width           =   195
      End
      Begin VB.Label xTotal_Net 
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   19
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label xTotal_Item 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   7
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   14
         Top             =   540
         Width           =   1860
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   810
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   7560
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
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
         Picture         =   "purchase.frx":1078B
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase.frx":1295B
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
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
         Picture         =   "purchase.frx":14AA3
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase.frx":16C6B
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
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
         Picture         =   "purchase.frx":18DBA
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase.frx":1AF9A
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   41
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
         Picture         =   "purchase.frx":1D0F5
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase.frx":1F2B1
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   0
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
End
Attribute VB_Name = "Purchasefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Integer
Dim con As New ADODB.Connection
Public bedit As Boolean, nRound As Integer, nRoundG As Integer
Dim CardTable As ADODB.Recordset, cFile As String, cFileHeader As String, cData2 As String
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchSup As New Search3
Public sDoc_no As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[CODE]", addstring(xCode.BoundText))
aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))
aInsert = AddFlag(aInsert, "[BOX]", addstring(xbox.BoundText))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.Text))
aInsert = AddFlag(aInsert, "[TAX1]", Val(xTax1.Text))
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert = AddFlag(aInsert, "[DOC_NO]", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, " DOC_NO = " & addstring(xDoc_No.Text))
End If
myreplaceGrd Row
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
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 1)
        grid1.TextMatrix(grid1.Row, 3) = 1
        Grid1_AfterEdit grid1.Row, grid1.Col
        If bNew Then
            grid1.Row = grid1.Rows - 1
        Else
            CellPos 13, grid1.Row, 1
        End If
        If grid1.Row < grid1.Rows - 1 Then oSearchItem.Hide
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    oSearchDoc.Hide
ElseIf ActiveControl.Name = xCode.Name Then
    ActiveControl.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
End Sub
Private Sub cmdClient_Click()
Dim sCode As String
sCode = xCode.BoundText
supfrm.Show 1
data2.Refresh
xCode.BoundText = sCode
If Not xCode.MatchedWithList Then xCode.BoundText = ""
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
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
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
Private Sub cmdPrint_Click()
doprint 1
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
    KeyCode = 0
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And grid1.Row = grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, 1) = "" Then
    KeyAscii = 0
    On Error Resume Next
    cmdSave.SetFocus
    Err.Clear
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 1 Then
    If Trim(.EditText) = "" Then
        Cancel = True
        Exit Sub
    End If
    grid1.EditText = grid1.EditText
    If Not validItem(.EditText, con) Then
        MsgBox "«·’‰› «·„”Ã· €Ì— „ÊÃÊœ"
        Cancel = True
        Exit Sub
    End If
End If
End With
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_Validate(Cancel As Boolean)
'With XCODE
'    If Not .MatchedWithList Then .BoundText = ""
'End With
End Sub

Private Sub xRate_LostFocus()
myLostFocus xRate
If Val(xRate.Text) <> 0 Then
    If Round(Val(xRate.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotal_Item.Caption) * 100, nRound) Or xDiscount.Locked Then
        xDiscount.Text = Round((Val(xRate.Text) * Val(xTotal_Item.Caption)) / 100, nRound)
    End If
Else
   xDiscount.Text = ""
End If
Calctotals
End Sub
Private Sub xTax1_LostFocus()
myLostFocus xTax1
Calctotals
End Sub
Private Sub xrate1_LostFocus()
myLostFocus xrate1
If Val(xTotal_Net.Caption) <> 0 Then
    If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xTax1.Locked Then
        xTax1.Text = Round((Val(xrate1.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
    End If
Else
   xTax1.Text = ""
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
xCode.SetFocus
End Sub
Private Sub cmdSave_Click()
mysave
If sDoc_no <> "" Then Unload Me
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
nRound = 2
openCon con
bedit = True
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
    lblBox.Visible = False
    xbox.Visible = False
End Select


data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

'data2.ConnectionString = strCon
'data2.RecordSource = "SELECT * FROM FILE4_10"
'Set XCODE.RowSource = data2
Set data2.Recordset = myRecordSet("SELECT * FROM FILE4_10", con)
Set xCode.RowSource = data2
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("Select * From file0_50 where code > '500000'", con)
Set xbox.RowSource = DATA3
xbox.ListField = "Desca"
xbox.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set Purchasefrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If grid1.Col = 1 Then
    GrdDesc Row
End If

If Not validRow(Row) Then Exit Sub
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
        'grid1.TextMatrix(Row, 7) = nBalance
    End If
End If
Calctotals
End With
End Sub
Private Sub Grid1_EnterCell()
If (grid1.Col = 2 Or grid1.Col = 5) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
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

If xCode.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
For i = 1 To grid1.Rows - 2
    If Not validRow(i) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xDoc_No.Text = CardTable!DOC_NO
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xStore.BoundText = CardTable!store & ""
xbox.BoundText = CardTable!BOX & ""
xCode.BoundText = CardTable!code & ""
xDiscount.Text = Myvalue(Val(CardTable!Discount & ""), "00.00")
xTax1.Text = Myvalue(Val(CardTable!TAX1 & ""), "00.00")
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Handlecontrols LoadMode
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
xDate.Text = Format(Date, "YYYY-MM-DD")
xStore.BoundText = sDefStore
xStore.Enabled = xStore.BoundText = ""
xbox.BoundText = ""
xBalance.Caption = ""
xCode.BoundText = ""
xDiscount.Text = ""
xTotal.Caption = ""
xTax1.Text = ""
xTotal_Item.Caption = ""
xDiscount.Text = ""
xRate.Text = ""
xTax1.Text = ""
xrate1.Text = ""
xtotalQuant.Caption = ""
grid1.Rows = 1
grid1.AddItem ""
Fixgrd
MakeSerial 1
Handlecontrols DefineMode
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdSave.Enabled = (bedit)
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xDoc_No.Tag = nMode
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
Calctotals
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
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRow(grid1.Row) Then
    grid1.AddItem "", grid1.Row
    MakeSerial grid1.Row - 1
ElseIf KeyCode = 112 Then
    ItemsLookupAll Me, oSearchItem
    ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bedit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
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
con.RollbackTrans
Err.Clear
End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 2) = ""
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
grid1.TextMatrix(Row, 1) = grid1.TextMatrix(Row, 1)
Dim aRet As Variant
aRet = ItemFields(grid1.TextMatrix(Row, 1), con)
If Not IsEmpty(aRet) Then
    grid1.TextMatrix(Row, 2) = retFlag(aRet, "DESCA")
End If
End Sub
Private Function Calctotals()
Dim nTotal As Double, nDiscount As Double, nTotal_item As Double, nTotalQuant As Double
With grid1
For i = 1 To grid1.Rows - 2
    .TextMatrix(i, 5) = Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4))
    nTotal_item = nTotal_item + Val(.TextMatrix(i, 5))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 3))
Next

xTotal_Item.Caption = nTotal_item
xTotal_Net.Caption = nTotal_item - Val(xDiscount.Text)

If Val(xTotal_Item.Caption) <> 0 Then
    If Round(Val(xRate.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotal_Item.Caption) * 100, nRound) Then
        xRate.Text = Round((Val(xDiscount.Text) / Val(xTotal_Item.Caption)) * 100, nRound)
    End If
Else
    xRate.Text = ""
End If


If Val(xTotal_Net.Caption) <> 0 Then
    If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Then
        xrate1.Text = Myvalue(Round((Val(xTax1.Text) / Val(xTotal_Net.Caption)) * 100, nRound))
    End If
Else
    xrate1.Text = ""
End If

'If Val(xTotal_Net.Caption) <> 0 Then
'    If Round(Val(xRate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xTax1.Locked Then
'        xTax1.Text = Round((Val(xRate1.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
'    End If
'Else
'    xTax1.Text = ""
'End If
xTotal.Caption = Val(xTotal_Item.Caption) + Val(xTax1.Text) - Val(xDiscount.Text)
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
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
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
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
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

For i = 1 To grid1.Rows - 2
    temptable.AddNew
    If myPublic = 0 Then
        temptable!str21 = "„” ‰œ „‘ —Ì«   "
    Else
        temptable!str21 = "„” ‰œ „—œÊœ „‘ —Ì«   "
    End If
    temptable!str1 = xDoc_No.Text
    temptable!str2 = xDate.Text
    temptable!str3 = Format(xCode.BoundText)
    temptable!str4 = xCode.Text
    temptable!str5 = xStore.Text
    temptable!Str11 = TurnValue(grid1.TextMatrix(i, 1))
    
    temptable!str12 = TurnValue(grid1.TextMatrix(i, 2))
    temptable!val1 = Val(grid1.TextMatrix(i, 3))
    temptable!val2 = Val(grid1.TextMatrix(i, 4))
    temptable!val4 = Val(grid1.TextMatrix(i, 5))
    temptable!val10 = i
    temptable!val11 = IIf(Not isOut, 0, 1)
    
       
  
    temptable!Val5 = Val(xTotal_Item.Caption)
    temptable!Val7 = Val(xDiscount.Text)
    temptable!Val8 = Val(xTotal.Caption)
    temptable!Val9 = Val(xtotalQuant.Caption)
    
    temptable!str10 = MyOnly(Val(xTotal.Caption))
    temptable!Val9 = myPublic
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If nPrint = 1 Then
    main.REPORT1.ReportFileName = App.Path & "\Reports\purchase.rpt"
Else
    main.REPORT1.ReportFileName = App.Path & "\Reports\purchase_A.rpt"
End If
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
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
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "Doc_no", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", Val(.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "PRICE", Val(.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "TOTAL", Val(.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(.TextMatrix(i, 6)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    myRemove grid1.Row
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
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
    data10.RecordSource = cString
    data10.Refresh
    grid1.AddItem ""
    MakeSerial
End With
Calctotals
Fixgrd
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 4 Then
    grid1.Col = Col + 1 + IIf(Col = 1, 1, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, IIf(grid1.TextMatrix(Row + 1, 1) <> "", 3, 1)
    grid1.ShowCell Row + 1, IIf(grid1.TextMatrix(Row + 1, 1) <> "", 3, 1)
Else
    grid1.Select Row, Col
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
Private Sub MyAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial
Calctotals
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
Private Sub xstore_GotFocus()
myGotFocus xStore
End Sub
Private Sub xStore_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStorexStore.BoundText = ""
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not xCode.MatchedWithList Then xCode.BoundText = ""
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xrate1_GotFocus()
myGotFocus xrate1
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xTax1_GotFocus()
myGotFocus xTax1
End Sub
