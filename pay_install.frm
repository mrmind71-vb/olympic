VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form pay_installfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄—Ê÷"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11655
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   7215
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
         Picture         =   "pay_install.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "pay_install.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "pay_install.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "pay_install.frx":6D48
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "pay_install.frx":95E2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "pay_install.frx":BB8E
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   675
      Width           =   7935
      Begin VB.TextBox Text1 
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
         Left            =   4995
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1260
         Width           =   1545
      End
      Begin VB.TextBox xPrice2 
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
         Left            =   4995
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox xPrice 
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
         Left            =   4995
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label4 
         Caption         =   "«·„»·€ «·„”œœ"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label xPrice_total1 
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
         Height          =   330
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   " «—ÌŒ «·”œ«œ"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·«” „«—…"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ «·„” ‰œ"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.CommandButton CMD_FIX 
      BackColor       =   &H00DEE7D3&
      Caption         =   "Fix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9360
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2160
      Top             =   -135
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   315
      Top             =   180
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   7950
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   20
      Top             =   7785
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   2430
      Top             =   -135
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Height          =   4740
      Left            =   180
      TabIndex        =   21
      Top             =   2385
      Width           =   11400
      _cx             =   20108
      _cy             =   8361
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
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7110
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
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
         Picture         =   "pay_install.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "pay_install.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   15
         TabStop         =   0   'False
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
         Picture         =   "pay_install.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "pay_install.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
         TabIndex        =   11
         TabStop         =   0   'False
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
         Picture         =   "pay_install.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "pay_install.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   12
         TabStop         =   0   'False
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
         Picture         =   "pay_install.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "pay_install.frx":1CE87
      End
   End
End
Attribute VB_Name = "pay_installfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim oSearch As New Search, osearchitem As New Search, oSearchItem2 As New Search, oSearch_Unit As New Search_empty, oSearchGroup As Search, oSearchGroup2 As New Search
Dim bEditRecord As Boolean, sFilter As String, bAct As Boolean, cFilter As String
Public nQuantFrm As Double
Public bedit As Boolean
Const LoadMode = 1, DefineMode = 2
Dim con As New ADODB.Connection
Dim CardTable As New ADODB.Recordset
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xdoc_no.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xDesca.SetFocus
    End If
    Err.Clear
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
Me.Top = 100
Me.Left = 100

openCon con

Set grid1.DataSource = data10

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xdoc_no.Text = RetZero(Newflag("FILE1_10", "ITEM"))
xDesca.SetFocus
End Sub
Private Sub CmdDel_Click()
Dim cString As String
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From FILE1_20  Where ITEM_MAIN = " & MyParn(xdoc_no.Text), nDelete
    con.Execute "Delete  From FILE1_10  Where ITEM = " & MyParn(xdoc_no.Text), nDelete
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "item < " & MyParn(xdoc_no.Text), , adSearchBackward, adBookmarkLast
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
SetKbLayout Lang_AR
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·’‰› »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
'RawLookupAll Me, oSearch, , "2", True, "up"
'ItemsLookupAll Me, oSearch, "FILE1_10.TYPE = '1'"
ItemsLookupAll Me, oSearch, cFilter, "1"
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
cmdAdd.Enabled = (nMode = LoadMode) And bEditRecord
cmdGroupAdd.Enabled = bEditRecord
CmdDel.Enabled = (nMode = LoadMode) And bEditRecord
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
cmdFilter.Visible = cmdFilter.Tag <> ""
xdoc_no.Enabled = Not (nMode = LoadMode)
xdoc_no.Tag = nMode
End Sub
Sub mydefine()
xdoc_no.Text = ""
xDesca.Text = ""
xGroup.BoundText = ""
xPrice.Text = ""
xPrice2.Text = ""
xPrice_total1.Caption = ""
xPrice_total2.Caption = ""
Handlecontrols DefineMode
StatusBar1.Panels(1).Text = ArbString("«÷«›… ”Ã· (" & CardTable.RecordCount + 1 & ")")
StatusBar1.Panels(2).Text = ""

Fixgrd
grid1.rows = 1
myAddItem

CellPos 13, grid1.rows - 2, grid1.Cols - 1
End Sub
Sub myload()
xdoc_no.Text = CardTable!doc_no
xDate.Text = myFormat_p(CardTable!Date)

xGroup.BoundText = CardTable!Group & ""
xPrice.Text = Myvalue(CardTable!PRICE)
xPrice2.Text = Myvalue(CardTable!Price2)
xWeight.Caption = Myvalue(CardTable!Weight)

myloadGrd

StatusBar1.Panels(1).Text = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
CellPos 13, grid1.rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Sub myproc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = FoundOtheritem(Row, Col, oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 0))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Exit Sub
    End If
    
    With grid1
    Dim bNew As Boolean
    bNew = .Row = grid1.rows - 1
   .TextMatrix(.Row, 0) = oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 0)
   .TextMatrix(.Row, 1) = oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 1)
   .TextMatrix(.Row, 2) = 1
   .TextMatrix(.Row, 3) = oSearchItem2.grid1.TextMatrix(oSearchItem2.grid1.Row, 2)
    grdDesc .TextMatrix(.Row, 0), .Row
    grid1_AfterEdit .Row, .Col
    'oSearchItem2.Hide
    If bNew Then
        CellPos 13, grid1.rows - 2, grid1.Cols - 1
    End If
    End With
ElseIf ActiveControl.Name = CmdInform.Name Then
    xdoc_no.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = cmdGroup.Name Then
    cmdGroup.Tag = oSearchGroup.grid1.TextMatrix(oSearchGroup.grid1.Row, 0)
    cmdGroup.Caption = oSearchGroup.grid1.TextMatrix(oSearchGroup.grid1.Row, 1)
    oSearchGroup.Hide
    openCardTable
    myUndo
ElseIf ActiveControl.Name = cmdGroupAdd.Name Then
    Dim sGroup As String
    sGroup = oSearchGroup2.grid1.TextMatrix(oSearchGroup2.grid1.Row, 0)
    oSearchGroup2.Hide
    nQuantFrm = 1
    Set quantFrm.myform = Me
    quantFrm.Show 1
    If nQuantFrm <> 0 Then
        If AddGroup(sGroup, nQuantFrm, chkBalAdd.Value = 1) Then
            openCardTable
            myUndo
            Inform " „  «÷«›… «’‰«› «·„Ã„Ê⁄… »‰Ã«Õ"
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SetKbLayout Lang_AR
'myDelFile TempSave(Me)
CardTable.Close
Set CardTable = Nothing
closeCon con
Unload oSearch
Set oSearch = Nothing

Err.Clear
End Sub
Private Sub grid1_EnterCell()
If (grid1.Col = 0 Or grid1.Col = 2) And bEditRecord Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.rows - 1 And cmdSave.Enabled Then
    If MsgBox("«·’‰› ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from FILE1_20 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 112 And grid1.Col = 0 And grid1.Row <> 0 Then
    ItemsLookupAll Me, oSearchItem2
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
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
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
'If Trim(xdoc_no.Text) = "" Then
'    MsgBox "ﬂÊœ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
'    Exit Function
'End If

If Trim(xDesca.Text) = "" Then
    MsgBox "≈”„ «·’‰› €Ì— „”Ã·"
    Exit Function
End If
        
If Not xGroup.MatchedWithList Then
    MsgBox "„Ã„Ê⁄… «·’‰› €Ì— „”Ã·…"
    Exit Function
End If

MYVALID = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not MYVALID(True) Then
    On Error Resume Next
    .SetFocus
    Err.Clear
    myloadGrd
    If Row < .rows - 1 Then
        .Select Row, Col
    Else
        CellPos 13, .rows - 2, .Cols - 1
    End If
    Exit Sub
End If

If Not validRow(Row) Then
    CalcTotals
    Exit Sub
End If

If Row = .rows - 1 Then myAddItem
CalcTotals
If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        xdoc_no.Tag = LoadMode
        xdoc_no.Enabled = False
    End If
    If .TextMatrix(Row, .Cols - 1) = "" Then
        myloadGrd
    End If
Else
    myloadGrd
End If
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
'If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
'    If Not validRow(OldRow) Then .RemoveItem OldRow
'End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .rows - 1 And .Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then .RemoveItem .Row
End With
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Not ValidNum(.TextMatrix(Row, 0), 6) Then Exit Function
If mRound(.TextMatrix(Row, 2)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If ValidNum(Format(.EditText)) Then .EditText = RetZero(.EditText)
    If Not ValidNum(.EditText, 6) Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    Else
        grid1.EditText = grid1.EditText
        If Not grdDesc(grid1.EditText, Row) Then
            MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ"
            Cancel = True
            Exit Sub
        End If
    End If
ElseIf Col = 2 Then
    If mRound(.EditText) = 0 Then
        MsgBox "ﬂ„Ì«  €Ì— „”Ã·…"
        Cancel = True
    End If
End If
End With
End Sub
Private Function grdDesc(sItem As String, Row As Long, Optional bSkip As Boolean = False) As Boolean
If Not ValidNum(sItem) Then Exit Function
aRet = ItemFields(sItem, con)
If IsEmpty(aRet) Then Exit Function
grid1.TextMatrix(Row, 1) = retFlag(aRet, "DESCA") & ""
grdDesc = True
End Function
Private Sub myloadGrd()
Dim cString As String
cString = "Select FILE1_20.ITEM,FILE1_10.DESCA,FILE1_20.QUANT,FILE1_10.PRICE,FILE1_10.PRICE * FILE1_20.QUANT,FILE1_10.PRICE2,FILE1_20.QUANT * PRICE2, FILE1_10.WEIGHT_UNIT ,FILE1_20.ID" & _
          " From FILE1_20 INNER JOIN FILE1_10 ON FILE1_20.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "FILE1_20.ITEM_MAIN = " & MyParn(xdoc_no.Text)
cString = cString & " Order by FILE1_20.ID"
Set data10.Recordset = myRecordSet(cString, con)
myAddItem
Fixgrd
CalcTotals
End Sub
Private Sub Fixgrd()
With grid1
    .FormatString = "«·ﬂÊœ|" & "«·’‰›|" & "«·ﬂ„Ì…|" & "”⁄— „” Â·ﬂ|" & "«·≈Ã„«·Ì|" & "”⁄— Ã„·…|" & "«·≈Ã„«·Ì|" & "«·Ê“‰|"
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    .ColWidth(2) = 800
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColHidden(.Cols - 1) = True
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Function myreplace(Optional Row As Long = -1) As Boolean
con.BeginTrans
On Error GoTo myerror
con.Execute retHeaderString
myreplacegrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function retHeaderString() As String
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "[GROUP]", addstring(xGroup.BoundText))
aInsert = AddFlag(aInsert, "[PRICE]", mRound(xPrice.Text))
aInsert = AddFlag(aInsert, "[PRICE2]", mRound(xPrice2.Text))
aInsert = AddFlag(aInsert, "[AMOUNT]", "1")
aInsert = AddFlag(aInsert, "[OFFER]", "1")
If xdoc_no.Tag = DefineMode Then
    xdoc_no.Text = RetZero(Newflag("FILE1_10", "ITEM"))
    aInsert = AddFlag(aInsert, "ITEM", addstring(xdoc_no.Text))
    retHeaderString = addInsert(aInsert, "FILE1_10")
Else
    retHeaderString = addUpdate(aInsert, "FILE1_10", "ITEM = " & xdoc_no.Text)
End If
End Function
Private Sub myreplacegrd(Optional Row As Long = -1)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .rows - 2, Row)
        aInsert = AddFlag(Empty, "ITEM_MAIN", addstring(xdoc_no.Text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "QUANT", mRound(.TextMatrix(i, 2)))
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_20")
        Else
            con.Execute addUpdate(aInsert, "FILE1_20", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If ValidNum(xdoc_no.Text, 6) Then
        CardTable.Find "ITEM = " & xdoc_no.Text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT FILE1_10.* FROM FILE1_10"
cFilter = ""
cFilter = "FILE1_10.OFFER = 1"
If cmdFilter.Tag <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE1_10.ITEM IN(" & cmdFilter.Tag & ")"
If cmdGroup.Tag <> "" Then cFilter = cFilter & turn(cFilter, " AND ") & "FILE1_10.[GROUP] = " & addvalue(cmdGroup.Tag)
If cFilter <> "" Then cString = cString & " WHERE " & cFilter
cString = cString & " ORDER BY FILE1_10.[item]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Select Row, Col + 1 + IIf(Col = 0, 1, 0)
ElseIf Row < grid1.rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    grid1.ShowCell Row + 1, 0
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
End Sub
Private Sub xQuant_Prp_GotFocus()
myGotFocus xQuant_Prp
End Sub
Private Sub xQuant_Prp_LostFocus()
myLostFocus xQuant_Prp
End Sub
Private Sub xUnit_GotFocus()
myGotFocus xUnit
End Sub
Private Sub xUnit_LostFocus()
myLostFocus xUnit
End Sub
Private Sub xPackage_GotFocus()
myGotFocus xPackage
End Sub
Private Sub xPackage_LostFocus()
myLostFocus xPackage
End Sub
Private Sub xUnit_small_GotFocus()
myGotFocus xUnit_small
End Sub
Private Sub xUnit_small_LostFocus()
myLostFocus xUnit_small
End Sub
Private Sub xCost_GotFocus()
myGotFocus xCost
End Sub
Private Sub xCost_LostFocus()
myLostFocus xCost
End Sub

Private Sub xdoc_no_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xdoc_no_LostFocus()
myLostFocus xdoc_no
If ValidNum(Format(xdoc_no.Text)) Then xdoc_no.Text = RetZero(xdoc_no.Text)
If Not ValidNum(xdoc_no.Text, 6) Then
     If xdoc_no.Tag = LoadMode Then
        mydefine
    Else
        xdoc_no.Text = ""
    End If
Else
    If (Not (CardTable.EOF Or CardTable.BOF)) And xdoc_no.Tag = LoadMode Then
        If CardTable!Item & "" = Trim(xdoc_no.Text) Then
            Exit Sub
        End If
    End If
    
    CardTable.Find "ITEM = " & MyParn(xdoc_no.Text), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        myload
    ElseIf xdoc_no.Tag = LoadMode Then
        mydefine
    Else
        xdoc_no.Text = ""
    End If
End If
End Sub
Private Sub xDesca_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xPrice_GotFocus()
myGotFocus xPrice
End Sub
Private Sub xPrice_LostFocus()
myLostFocus xPrice
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xGroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xType_store_GotFocus()
myGotFocus xType_store
End Sub
Private Sub xType_Store_LostFocus()
myLostFocus xType_store
If Not xType_store.MatchedWithList Then xType_store.BoundText = ""
End Sub
Private Sub CalcTotals()
Dim nTotal As Double, nTotal2 As Double, i As Long, nQuant As Long
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, 4) = Myvalue((mRound(grid1.TextMatrix(i, 2)) * mRound(grid1.TextMatrix(i, 3))))
    nTotal = nTotal + (mRound(grid1.TextMatrix(i, 2)) * mRound(grid1.TextMatrix(i, 3)))
    

    grid1.TextMatrix(i, 6) = Myvalue((mRound(grid1.TextMatrix(i, 2)) * mRound(grid1.TextMatrix(i, 5))))
    nTotal2 = nTotal2 + (mRound(grid1.TextMatrix(i, 2)) * mRound(grid1.TextMatrix(i, 5)))
    nQuant = nQuant + mRound(grid1.TextMatrix(i, 2))
Next
xPrice_total1.Caption = Myvalue(nTotal)
xPrice_total2.Caption = Myvalue(nTotal2)
StatusBar1.Panels(2).Text = "⁄œœ «’‰«› «·⁄—÷ : " & Myvalue(grid1.rows - 2)
StatusBar1.Panels(3).Text = "ﬂ„Ì«  «’‰«› «·⁄—÷ : " & nQuant
End Sub
