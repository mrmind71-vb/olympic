VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì— „»Ì⁄« "
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18870
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
   ScaleHeight     =   8280
   ScaleWidth      =   18870
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "÷»ÿ «·„»Ì⁄« "
      Height          =   645
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   1305
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   0
      Width           =   6810
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "  »œÊ‰"
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
         Height          =   255
         Index           =   0
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ÃÂ…"
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
         Height          =   255
         Index           =   3
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„ «·»Ê·Ì’…"
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
         Height          =   255
         Index           =   2
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   5265
         Picture         =   "Sales.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   135
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1215
      Width           =   2040
      Begin VB.CommandButton cmdAddTravel 
         Caption         =   "«÷«›… »Ê«·’ «·‘Õ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   135
         Width           =   1950
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   13365
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":242A
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
         Picture         =   "Sales.frx":4848
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
         Picture         =   "Sales.frx":70E2
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
         Picture         =   "Sales.frx":968E
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   8595
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   810
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":BE61
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1320
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
         Picture         =   "Sales.frx":E3DA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   675
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
      Left            =   -1395
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
      Left            =   0
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
      Left            =   2205
      Top             =   405
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
      Left            =   1485
      Top             =   -180
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   -1890
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
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   2025
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1215
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   90
      TabIndex        =   5
      Top             =   1980
      Width           =   18690
      _cx             =   32967
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   10035
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   630
      Width           =   8745
      Begin VB.TextBox xInv_no 
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
         Left            =   3510
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox xNotes 
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
         Left            =   90
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   7620
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
         Left            =   6615
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox xCode 
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
         Left            =   6615
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
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
         TabIndex        =   2
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "„·ÕÊŸ…"
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
         Left            =   7785
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   945
         Width           =   615
      End
      Begin VB.Label xCode_Desca 
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
         Left            =   3510
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
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
         Left            =   7785
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   840
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
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
         Left            =   7785
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   480
      End
   End
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   0
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
   Begin VB.Frame Frame6 
      Height          =   1365
      Left            =   15075
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   6705
      Width           =   3705
      Begin VB.TextBox xDiscount 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   1035
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   585
         Width           =   1185
      End
      Begin VB.TextBox xRate 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   405
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
         Width           =   600
      End
      Begin VB.Label xTotal_Net 
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
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label11 
         Caption         =   "«·’«›Ì"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   945
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„"
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
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   375
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
         Height          =   375
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "≈Ã„«·Ì «·»Ê«·’"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   1290
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1320
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   6705
      Width           =   9150
      Begin VB.TextBox xWeight 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   1440
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox xrate1 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   5895
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   690
      End
      Begin VB.TextBox xTax1 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   6615
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox xrate2 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   720
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         Width           =   690
      End
      Begin VB.TextBox xtax2 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   1440
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label xTotal 
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   900
         Width           =   1860
      End
      Begin VB.Label Label8 
         Caption         =   "«·≈Ã„«·Ì "
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
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "€—«„«  Ê„Ê«“Ì‰"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "÷—»Ì… „»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "÷—»Ì… Œ’„ Ê≈÷«›…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "ﬁÌ„… «·›« Ê—…"
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
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label xTotal1 
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   540
         Width           =   1860
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   6705
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   28
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
         Picture         =   "Sales.frx":1073D
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":1290D
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   29
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
         Picture         =   "Sales.frx":14A55
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":16C1D
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   30
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
         Picture         =   "Sales.frx":18D6C
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":1AF4C
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   31
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
         Picture         =   "Sales.frx":1D0A7
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":1F263
      End
   End
End
Attribute VB_Name = "salesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Dim nRate1 As Long, nRate2 As Long, bActivated As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String
Dim cFilter As String
Public bRetvalue As Boolean
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchClient As New Search3, oSearch_Travel As New Search3
Dim oTravel_Add As New travel_addfrm
Dim bedit As Boolean
Dim nRound As Integer
Dim con As New ADODB.Connection
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CODE", addstring(xCode.Text))
aInsert = AddFlag(aInsert, "[INV_NO]", addstring(xInv_no.Text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.Text))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.Text))
aInsert = AddFlag(aInsert, "[RATE]", Val(xRate.Text))
aInsert = AddFlag(aInsert, "[TAX1]", Val(xTax1.Text))
aInsert = AddFlag(aInsert, "[RATE1]", Val(xrate1.Text))
aInsert = AddFlag(aInsert, "[TAX2]", Val(xtax2.Text))
aInsert = AddFlag(aInsert, "[RATE2]", Val(xrate2.Text))
aInsert = AddFlag(aInsert, "[WEIGHT]", Val(xWeight.Text))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(sUserName))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Text = "" Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE6_20H", "doc_no")))
    aInsert = AddFlag(aInsert, "[DOC_NO]", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    con.Execute addUpdate(aInsert, "FILE6_20H", "doc_no = " & addstring(xDoc_No.Text))
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
    nFound = grid1.FindRow(oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0), , 1)
    If nFound <> -1 Then
         MsgBox "»Ê·Ì’… « ·‘Õ‰ „ÊÃÊœ… ›Ï «·”ÿ— " & nFound
         Exit Sub
    End If
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 2)
    grid1.TextMatrix(grid1.Row, 3) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 3)
    grid1.TextMatrix(grid1.Row, 4) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 4)
    grid1.TextMatrix(grid1.Row, 5) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 5)
    grid1.TextMatrix(grid1.Row, 6) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 6)
    grid1.TextMatrix(grid1.Row, 7) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 11)
    grid1.TextMatrix(grid1.Row, 8) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 12)
    grid1.TextMatrix(grid1.Row, 9) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 13)
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearch_Travel
        CellPos 13, grid1.Row, 2
    Else
        grid1.Select grid1.Rows - 1, 2
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    Unload oSearchDoc
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    xCode_Validate False
    Unload oSearchClient
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdAddTravel_Click()
If Not MYVALID(True) Then Exit Sub
oTravel_Add.sCode = xCode.Text
oTravel_Add.sCode_Desca = xcode_Desca.Caption
Set oTravel_Add.myForm = Me
oTravel_Add.Show 1
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_No.Text)
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
Private Sub CmdInform_Click()
CardLookup cFilter
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
xInv_no.SetFocus
On Error Resume Next
'grid1.SetFocus
Err.Clear
End Sub

Private Sub cmdPrint_Click()
n = RetIndex(Option1)
doprint xDoc_No.Text, RetIndex(Option1)
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not CheckRows Then Exit Sub
mysave
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub

Private Sub Command1_Click()

Dim loctable As New ADODB.Recordset
loctable.Open "select * FROM TRAVEL_H WHERE WEIGHT_TOTAL <> 0 AND CLASS = 1", con, adOpenSt, adLockReadOnly, adCmdText
Dim aInsert As Variant
Me.Caption = "0"
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "CLASS", Val(loctable!Weight & ""))
    aInsert = AddFlag(aInsert, "WEIGHT", Val(loctable!Class & ""))
    con.Execute addUpdate(aInsert, "TRAVEL_H", "DOC_NO = " & MyParn(loctable!DOC_NO))
    loctable.MoveNext
    Me.Caption = Val(Me.Caption) + 1
Loop
loctable.Close

loctable.Open "select FILE6_20.TRAVEL,TRAVEL_H.CLASS,TRAVEL_H.WEIGHT,TRAVEL_H.TOTAL,FILE6_20.ID FROM FILE6_20 INNER JOIN TRAVEL_H ON FILE6_20.TRAVEL = TRAVEL_H.DOC_NO", con, adOpenSt, adLockReadOnly, adCmdText
Me.Caption = "0"
Do Until loctable.EOF
    aInsert = AddFlag(Empty, "QUANT", Val(loctable!Weight & ""))
    aInsert = AddFlag(aInsert, "PRICE", Val(loctable!Class & ""))
    aInsert = AddFlag(aInsert, "TOTAL", Val(loctable!TOTAL & ""))
    con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & loctable!ID)
    loctable.MoveNext
    Me.Caption = Val(Me.Caption) + 1
Loop
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Not bActivated Then
    bActivated = True
    If xDoc_No.Tag = LoadMode Then
        grid1.SetFocus
        Err.Clear
    Else
        xInv_no.SetFocus
    End If
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1_Validate False
    cmdSave_Click
    KeyCode = 0
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
nRound = 2
openCon con

Set grid1.DataSource = data11
data11.ConnectionString = strCon

Fixgrd
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing

closeCon con

Unload oSearchDoc
Unload oSearchClient
Unload oSearch_Travel
Set salesfrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1

Calctotals
If Not validRow(Row) Then Exit Sub

If Row = .Rows - 1 Then MyAddItem
Calctotals
    
If myreplace(Row) Then
    HandleCntEdit
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
        myloadgrd
        grid1.Row = grid1.Rows - 1
        grid1.ShowCell grid1.Rows - 1, 1
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
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Trim(xCode.Text) = "" Then Exit Sub
Travel_LookupAll Me, oSearch_Travel, "((NOT POLICY IS NULL) AND (NOT DATE_POLICY IS NULL) AND TRAVEL_H.CODE = " & MyParn(xCode.Text) & ")", True
End Sub
Private Sub Grid1_EnterCell()
If (grid1.Col = 2) And bedit Then
    grid1.Editable = flexEDKbdMouse
Else
   grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
'If grid1.Rows < 1 Then Exit Sub
'If grid1.Row = 0 Then
'    grid1.Row = 1
'    grid1.Col = 1
'End If
Grid1_EnterCell
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ClientLookupAll Me, oSearchClient
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xCode_Validate(Cancel As Boolean)
xcode_Desca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file3_10 where code = " & MyParn(xCode.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xcode_Desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Sub xDate_delivery_LostFocus()
myValidDate xDate_delivery
End Sub
Private Sub xDate_delivery_v_LostFocus()
myValidDate xDate_delivery_v
End Sub
Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
Calctotals
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
Dim i As Integer

If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xcode_Desca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If
With grid1
End With
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xDoc_No.Text = CardTable!DOC_NO
xInv_no.Text = CardTable!inv_no & ""
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xNotes.Text = CardTable!Notes & ""
xCode.Text = CardTable!code & ""
xcode_Desca.Caption = CardTable!ClientDesca & ""
xDiscount.Text = Myvalue(CardTable!Discount)
xrate1.Text = Myvalue(CardTable!RATE1)
xrate2.Text = Myvalue(CardTable!RATE2)
xTax1.Text = Myvalue(CardTable!TAX1)
xtax2.Text = Myvalue(CardTable!TAX2)
xWeight.Text = Myvalue(CardTable!Weight)
myloadgrd
Handlecontrols LoadMode
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Sub mydefine()
xDoc_No.Text = ""
xInv_no.Text = ""
xDate.Text = Format(Date, "YYYY-MM-DD")
xCode.Text = ""
xcode_Desca.Caption = ""
xDiscount.Text = ""
xTotal.Caption = ""
xTotal_Net.Caption = ""
xTotal_Item.Caption = ""
xDiscount.Text = ""
xRate.Text = ""
xTax1.Text = ""
xrate1.Text = Myvalue(nRate1)
xtax2.Text = ""
xrate2.Text = Myvalue(nRate2)
xNotes.Text = ""
xWeight.Text = ""
grid1.Rows = 1
MyAddItem
Fixgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xDoc_No.Enabled = (nMode = DefineMode)
xCode.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
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
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bedit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM FILE6_20 WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        myRemove grid1.Row
    End If
ElseIf eyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub GrdDesc(Row)
If Trim(grid1.TextMa1trix(Row, 1)) = "" Then Exit Sub
grid1.TextMatrix(Row, 2) = ""
grid1.TextMatrix(Row, 3) = ""
grid1.TextMatrix(Row, 4) = ""
grid1.TextMatrix(Row, 5) = ""
grid1.TextMatrix(Row, 6) = ""

Dim aRet As Variant
aRet = ItemFields(grid1.TextMatrix(Row, 1), con)
If Not IsEmpty(aRet) Then
    grid1.TextMatrix(Row, 2) = retFlag(aRet, "DESCA")
    If grid1.TextMatrix(Row, 3) = "" Then grid1.TextMatrix(Row, 3) = "1"
    If Trim(grid1.TextMatrix(Row, 4)) = "" Then grid1.TextMatrix(Row, 4) = Val(retFlag(aRet, "price"))
    grid1.TextMatrix(Row, 6) = Val(retFlag(aRet, "price"))
End If
End Sub
Private Function Calctotals(Optional nMode As Integer = 0)
Dim nTotal_item As Double, nTotat_Exam As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal_item = nTotal_item + Val(.TextMatrix(i, 9))
Next
End With
xTotal_Item.Caption = nTotal_item
If Val(xTotal_Item.Caption) <> 0 Then
    If Round(Val(xRate.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotal_Item.Caption) * 100, nRound) Then
        xRate.Text = Myvalue(Round((Val(xDiscount.Text) / Val(xTotal_Item.Caption)) * 100, nRound))
    End If
Else
    xRate.Text = ""
End If
xTotal_Net.Caption = Round(Val(xTotal_Item.Caption) - Val(xDiscount.Text), 2)

If Val(xTotal_Net.Caption) <> 0 Then
'    If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Then
'        xrate1.Text = Myvalue(Round((Val(xTax1.Text) / Val(xTotal_Net.Caption)) * 100, nRound))
'    End If
    If Val(xrate1.Text) <> 0 Then
        If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xTax1.Locked Then
            xTax1.Text = Round((Val(xrate1.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
        End If
    Else
       xTax1.Text = ""
    End If
Else
    xTax1.Text = ""
End If

If Val(xTotal_Net.Caption) <> 0 Then
    If Round(Val(xrate2.Text), nRound) <> Round(Val(xtax2.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xtax2.Locked Then
        xtax2.Text = Round((Val(xrate2.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
    End If
'    If Round(Val(xrate2.Text), nRound) <> Round(Val(xtax2.Text) / Val(xTotal_Net.Caption) * 100, nRound) Then
'        xrate2.Text = Myvalue(Round((Val(xtax2.Text) / Val(xTotal_Net.Caption)) * 100, nRound))
'    End If
Else
    xtax2.Text = ""
End If
xTotal1.Caption = Val(xTotal_Net.Caption) + Val(xTax1.Text) - Val(xtax2.Text)
xTotal.Caption = Val(xTotal1.Caption) + Val(xWeight.Text)
End Function
Private Sub CardLookup(Optional pWhere As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,INV_NO,Convert(VARCHAR(10),[DATE],111),FILE3_10.Desca " & _
                  " FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE"
If pWhere <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & pWhere
End If
Generalarray(2) = "Order by Date , doc_no "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·⁄„Ì·-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 or %%FILE3_10.DESCA%% OR " & _
                  "##date##)"

listarray(1, 0) = "—ﬁ„ «·›« Ê—…"
listarray(1, 1) = "(%%INV_NO%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "—ﬁ„ «·›« Ê—…"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1400

GrdArray(3, 0) = "≈”„ «·⁄„Ì·"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub

Private Sub xInv_no_GotFocus()
myGotFocus xInv_no
End Sub
Private Sub xInv_no_LostFocus()
myLostFocus xInv_no
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
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "»Ê·Ì’… «·‘Õ‰|" & "«· «—ÌŒ|" & "«·»Ì«‰|" & " Õ„Ì·|" & " ⁄ Ìﬁ|" & "«·Ê“‰|" & "«·›∆…|" & "«·‰Ê·Ê‰|" & "«·Õ„Ê·…|"
.ColWidth(0) = 500
.ColWidth(1) = 1000
.ColWidth(2) = 1200
.ColWidth(3) = 1500
.ColWidth(4) = 2000
.ColWidth(5) = 2000
.ColWidth(6) = 2000
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200

.ColHidden(1) = True
.ColComboList(2) = "..."
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xRateDis_LostFocus()
xDiscount.Text = Fix((Val(xTotal_Item.Caption) * Val(xRateDis.Text) / 100))
End Sub
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "TRAVEL", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "QUANT", Val(grid1.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "PRICE", Val(grid1.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "TOTAL", Val(grid1.TextMatrix(i, 9)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE6_20.TRAVEL,TRAVEL_H.POLICY,Convert(VARCHAR(10),[DATE_POLICY],111),TRAVEL_H.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,FILE6_20.QUANT,FILE6_20.PRICE, FILE6_20.TOTAL,CARGO_CODES.DESCA,FILE6_20.ID " & _
              " FROM FILE6_20 INNER JOIN TRAVEL_H ON FILE6_20.TRAVEL = TRAVEL_H.DOC_NO" & _
              " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
              " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE " & _
              " LEFT JOIN CARGO_CODES ON TRAVEL_H.CARGO = CARGO_CODES.CODE"
    cString = cString & turn(cString) & " FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
    data11.RecordSource = cString
    data11.Refresh
    MyAddItem
End With
Calctotals
Fixgrd
End Sub
Private Function mysave() As Boolean
If Not MYVALID Then Exit Function
Calctotals
If Not myreplace Then Exit Function
Inform " „ Õ›Ÿ «·„” ‰œ"
openCardTable
myUndo
End Function
Private Function doprint(sDoc_no, nOption As Integer) As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Dim loctable As New ADODB.Recordset

cString = "SELECT FILE6_20.DOC_NO,FILE3_10.DESCA AS CODE_DESCA,FILE6_20.TRAVEL,TRAVEL_H.POLICY,TRAVEL_H.DATE AS DATE_TRAVEL,FILE6_20H.DISCOUNT,FILE6_20H.RATE,FILE6_20H.TAX1,FILE6_20H.RATE1,FILE6_20H.TAX2,FILE6_20H.RATE2,FILE6_20H.WEIGHT AS WEIGHT_CHARGE,FILE6_20H.DATE,TRAVEL_H.DESCA,PLACE_CODES.DESCA AS PLACE1,PLACE_CODES_B.DESCA AS PLACE2,FILE6_20.QUANT,FILE6_20.PRICE, FILE6_20.TOTAL,CARGO_CODES.DESCA AS CARGO_DESCA,FILE6_20.ID " & _
          " FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
          " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE " & _
          " INNER JOIN TRAVEL_H ON FILE6_20.TRAVEL = TRAVEL_H.DOC_NO" & _
          " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
          " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE " & _
          " LEFT JOIN CARGO_CODES ON TRAVEL_H.CARGO = CARGO_CODES.CODE"
cString = cString & turn(cString) & " FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
If nOption = 0 Then
    cString = cString & " ORDER BY FILE6_20.ID"
ElseIf nOption = 1 Then
    cString = cString & " ORDER BY TRAVEL_H.DATE"
ElseIf nOption = 2 Then
    cString = cString & " ORDER BY CONVERT(DECIMAL,TRAVEL_H.POLICY,2)"
ElseIf nOption = 3 Then
    cString = cString & " ORDER BY PLACE_CODES.DESCA"
End If

'Dim nTotal As Double
'nTotal = Val(GetField("select sum(file6_20.total) from file6_20 where doc_no = " & MyParn(sDoc_no)) & "")
Dim i As Long
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    i = i + 1
    temptable.AddNew
    temptable!val1 = loctable!TOTAL
    temptable!val2 = loctable!price
    temptable!Val3 = loctable!Quant
    temptable!str1 = loctable!cargo_Desca
    temptable!str2 = loctable!place1
    temptable!str3 = loctable!place2
    temptable!str4 = loctable!date_Travel
    temptable!str5 = loctable!Policy
    temptable!Val5 = loctable!Discount
    temptable!Val7 = loctable!TAX1
    temptable!Val8 = loctable!RATE1
    temptable!val9 = loctable!TAX2
    temptable!val10 = loctable!RATE2
    temptable!val11 = loctable!WEIGHT_CHARGE
    temptable!str21 = retFlag(aAddress, "DESCA")
    
    temptable!str10 = ArbString(Val(sDoc_no))
    temptable!Str11 = loctable!code_Desca
    temptable!str12 = TurnValue(Format(loctable!Date, "YYYY-MM-DD"))
    temptable!Val15 = i
    temptable.Update
    loctable.MoveNext
Loop

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.ReportFileName = App.Path & "\Reports\INVOICE.rpt"
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
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
xCode.Enabled = False
cmdSave.Enabled = (bedit) And grid1.Rows > 2
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE6_20H.*,FILE3_10.DESCA AS CLIENTDESCA,FILE3_10.PRICE,FILE3_10.CASH AS ISCASH FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.Code = FILE3_10.CODE"
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
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    grid1.Col = Col + 1 + IIf(Col = 1 Or Col = 4, 1, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, IIf(grid1.TextMatrix(Row + 1, 1) <> "", 3, 2)
    grid1.ShowCell Row + 1, 2
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub MyAddItem()
grid1.AddItem ""
MakeSerial
End Sub
Private Function CheckRows() As Boolean
If grid1.Rows < 3 Then Exit Function
CheckRows = True
End Function
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
Calctotals
MakeSerial
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 7 Then
    If Not IsNumeric(.EditText) <> "" And Trim(.EditText) <> "" Then
        Cancel = True
    End If
End If
End With
End Sub

Private Sub xrate1_LostFocus()
myLostFocus xrate1
'If Val(xrate1.Text) <> 0 Then
'    If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xTax1.Locked Then
'        xTax1.Text = Round((Val(xrate1.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
'    End If
'Else
'   xTax1.Text = ""
'End If
Calctotals
End Sub
Private Sub xrate2_LostFocus()
myLostFocus xrate2
'If Val(xrate2.Text) <> 0 Then
'    If Round(Val(xrate2.Text), nRound) <> Round(Val(xtax2.Text) / Val(xTotal_Net.Caption) * 100, nRound) Or xtax2.Locked Then
'        xtax2.Text = Round((Val(xrate2.Text) * Val(xTotal_Net.Caption)) / 100, nRound)
'    End If
'Else
'   xtax2.Text = ""
'End If
Calctotals
End Sub
Private Sub xTax1_LostFocus()
myLostFocus xTax1
If Val(xTotal_Net.Caption) <> 0 Then
    If Round(Val(xrate1.Text), nRound) <> Round(Val(xTax1.Text) / Val(xTotal_Net.Caption) * 100, nRound) Then
        xrate1.Text = Myvalue(Round((Val(xTax1.Text) / Val(xTotal_Net.Caption)) * 100, nRound))
    End If
Else
    xrate1.Text = ""
End If
Calctotals
End Sub
Private Sub xtax2_LostFocus()
myLostFocus xtax2
If Val(xTotal_Net.Caption) <> 0 Then
    If Round(Val(xrate2.Text), nRound) <> Round(Val(xtax2.Text) / Val(xTotal_Net.Caption) * 100, nRound) Then
        xrate2.Text = Myvalue(Round((Val(xtax2.Text) / Val(xTotal_Net.Caption)) * 100, nRound))
    End If
Else
    xrate2.Text = ""
End If
Calctotals
End Sub
Public Sub Addproc()
With oTravel_Add.grid1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, .Cols - 1)) <> 0 Then
            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
            grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(i, 0)
            grid1.TextMatrix(grid1.Rows - 1, 2) = .TextMatrix(i, 2)
            grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(i, 3)
            grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(i, 4)
            grid1.TextMatrix(grid1.Rows - 1, 5) = .TextMatrix(i, 5)
            grid1.TextMatrix(grid1.Rows - 1, 6) = .TextMatrix(i, 6)
            grid1.TextMatrix(grid1.Rows - 1, 7) = .TextMatrix(i, 7)
            grid1.TextMatrix(grid1.Rows - 1, 8) = .TextMatrix(i, 8)
            grid1.TextMatrix(grid1.Rows - 1, 9) = .TextMatrix(i, 9)
            grid1.AddItem ""
        End If
    Next
End With
mysave
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xrate1_GotFocus()
myGotFocus xrate1
End Sub
Private Sub xTax1_GotFocus()
myGotFocus xTax1
End Sub
Private Sub xrate2_GotFocus()
myGotFocus xrate2
End Sub
Private Sub xtax2_GotFocus()
myGotFocus xtax2
End Sub
Private Sub Xweight_GotFocus()
myGotFocus xWeight
End Sub
Private Sub Xweight_LostFocus()
myLostFocus xWeight
Calctotals
End Sub
