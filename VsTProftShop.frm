VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTProftShop 
   Caption         =   "„ «»⁄…  ›’Ì·Ï „»Ì⁄«  - √—’œ… _ —»Õ «·„Õ· ( «·”·«„ ﬂ·«” )"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_pr 
      Caption         =   "ÿ»«⁄… «’‰«› ·Ì” ·Â« —’Ìœ ›Ï «·„Õ· Ê ·Â« —’Ìœ »«·‘—ﬂ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   90
      Width           =   3705
   End
   Begin Threed.SSCommand xcount 
      Height          =   390
      Left            =   60
      TabIndex        =   26
      Top             =   825
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
      ButtonStyle     =   3
   End
   Begin VB.CommandButton cmd_bal 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄—÷ √—’œ… «·‘—ﬂ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1275
      Width           =   2175
   End
   Begin VB.CheckBox xprint 
      Alignment       =   1  'Right Justify
      Caption         =   "ÿ»«⁄… «·ﬂ·"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1275
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1695
      Width           =   1275
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1695
      Width           =   1275
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "⁄—÷"
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
      Left            =   2625
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1695
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.CheckBox xcost 
         Alignment       =   1  'Right Justify
         Caption         =   "”⁄— «· ﬂ·›… ··«’‰«› «· Ï ·Ì” ·Â«  ﬂ·›… ÂÊ ”⁄— «·Ã„·…"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   975
         Value           =   1  'Checked
         Width           =   4365
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7140
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1725
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox XDATEIMP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1725
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   590
         Width           =   4440
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1353
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   982
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   6000
         TabIndex        =   12
         Top             =   611
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XFACTNAME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1725
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” ‰œ"
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·—”«·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5805
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   17
         Top             =   640
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1005
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9555
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -135
      Top             =   300
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Bindings        =   "VsTProftShop.frx":0000
      Height          =   7335
      Left            =   150
      TabIndex        =   9
      Top             =   2250
      Width           =   14865
      _cx             =   26220
      _cy             =   12938
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14220542
      ForeColorSel    =   64
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoSearch      =   1
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   150
      Top             =   225
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   375
      Top             =   300
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
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   1500
      Top             =   75
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
End
Attribute VB_Name = "VsTProftShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConShop As New ADODB.Connection
Dim con As New ADODB.Connection
Dim MyData As String
Dim LastSalTable As New ADODB.Recordset
Dim ItemTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim MyBalItem As New ADODB.Recordset
Private Sub cmd_bal_Click()
With grid1
    For I = 2 To .Rows - 1
        MyBalItem.Filter = " ITEM = " & MyParn(.TextMatrix(I, 0))
        If Not MyBalItem.EOF Then
            .TextMatrix(I, 21) = Format(Val(MyBalItem!BAL & ""), "#0")
            cmd_bal.Caption = "⁄—÷ √—’œ… «·‘—ﬂ… " & Format(I - 1, "#0")
        End If
        If (Val(.TextMatrix(I, 20)) >= Val(.TextMatrix(I, 19))) And Val(.TextMatrix(I, 21)) > 0 Then
            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &HC0C0FF
        End If
    
    Next I
End With
End Sub
Private Sub cmd_pr_Click()
    DoPrintItemShop
End Sub

Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰  ›’Ì·Ï „Êﬁ› «’‰«› „Õ· «·”·«„ ﬂ·«” "
    cHead2 = " „‰  «—ÌŒ " & Format(xdate1.Text, "YYYY-MM-DD") & " ≈·Ï  «—ÌŒ " & Format(XDATE2.Text, "YYYY-MM-DD")
    With grid1
    If bopt1 Then
        .ColHidden(.Cols - 1) = True
        For I = 2 To .Rows - 1
            .RowHidden(I) = Not IIf(TurnValue(.TextMatrix(I, .Cols - 1), "", 0) = 0, False, True)
        Next I
        Load PrintGrd
        PrintGrd.doprint Me.grid1, 0.7, -2, cHead1, cHead2, , False, True, 7
        PrintGrd.Show 1
        .ColHidden(.Cols - 1) = False
        For I = 2 To .Rows - 1
            .RowHidden(I) = False
        Next I
    Else
        .ColHidden(.Cols - 1) = True
'       .ColHidden(5) = True
'       .ColHidden(7) = True
'       .ColHidden(10) = True
'       .ColHidden(16) = True
        
        
        For I = 2 To .Rows - 1
            .RowHidden(I) = Not IIf(TurnValue(.TextMatrix(I, .Cols - 1), "", 0) = 0, False, True)
        Next I
        Load PrintGrd
        PrintGrd.doprint Me.grid1, 0.85, -2, cHead1, cHead2, , False, False, 8
        PrintGrd.Show 1
        .ColHidden(.Cols - 1) = False
        
        .ColHidden(5) = False
'       .ColHidden(7) = False
'       .ColHidden(10) = False
'       .ColHidden(16) = False
        
        For I = 2 To .Rows - 1
            .RowHidden(I) = False
        Next I
    End If
    End With
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    cPathShop2 = GetDesca("select path from path")

    ConShop.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cPathShop2 & "\data.mdb"
    
    xdate1.Text = "3-6-2009"
    XDATE2.Text = Format(Date, "YYYY-MM-DD")
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
    Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    MyBalItem.Open "ITEM_BAL", con, adOpenStatic, adLockReadOnly, adCmdTable
    Set grid1.DataSource = data4
    data4.ConnectionString = ConShop.ConnectionString
    grid1.Rows = 1
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim MyData As String
MyData = App.Path & "\DATA\DATA.MDB"

cwhere = " date < " & DateSq(xdate1.Text)
cField5 = myiif(cwhere, "[IN]- [OUT] ") & " AS F_BAL"

cwhere = " date <= " & DateSq(XDATE2.Text)
cField17 = myiif(cwhere, "[IN]- [OUT] ") & " AS C_BAL"

cwhere = " date < " & DateSq(xdate1.Text)
cField6 = myiif(cwhere, "Val((FILE1_11.[in] - FILE1_11.[out] ) & '')* Val(FILE1_10.cost0 & '') ") & " AS cost_fbal"

cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '2' )"
cField71 = myiif(cwhere, "[IN] ") & " AS T_PURCH"

cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '7'  )"
cField72 = myiif(cwhere, " [OUT]") & " AS T_RPURCH"

cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '2' OR TYPE = '7' )"
cField73 = myiif(cwhere, "[IN] - [OUT]") & " AS T_NETPURCH"


cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '2' OR TYPE = '7' )"
cField8 = myiif(cwhere, "Val(( FILE1_11.IN - FILE1_11.[OUT] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))") & " AS TV_PURCH"

cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField10 = myiif(cwhere, "[OUT] - [IN]") & " AS T_SALES"


cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField11 = myiif(cwhere, "Val(( FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))") & " AS TV_SALES"

'myitemcost.costitem
cwhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(XDATE2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField12 = myiif(cwhere, "Val((FILE1_11.[out] - FILE1_11.[in] ) & '')* Val(FILE1_10.cost0 & '') ") & " AS Tcost_sales"


With grid1
'                           0                    1       2                 3                                 4
    cStrAll = " select  file1_10.item , file1_10.desca , FILE1_10.cost0  , FILE1_10.PRICE  , myitemcost.costitem , " & _
                cField5 & " , " & cField6 & " , " & cField71 & " , " & cField72 & " , " & cField73 & " , " & cField8 & " , ' ' as n9 ,  " & cField10 & " , " & cField11 & " , " & cField12 & " , ' ' AS N13 , ' ' AS N14 , ' ' AS N15 , ' ' AS N16 , " & cField17 & " , ' ' AS N20  " & _
                " FROM (((FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) LEFT JOIN file1_50 ON FILE1_10.[GROUP] = file1_50.CODE) LEFT JOIN FILE3_10 ON FILE1_11.CODECUST = FILE3_10.CODE  ) inner join myitemcost on file1_10.item = myitemcost.item "
    If xGroup.BoundText <> "" Then cStrAll = cStrAll & turn(cStrAll, " where ") & " file1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStrAll = cStrAll & turn(cStrAll, " where ") & " file1_50.[GROUP]   = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStrAll = cStrAll & turn(cStrAll, " where ") & " [Section] = " & xSection.BoundText
    If xDesca.Text <> "" Then cStrAll = cStrAll & turn(cStrAll, " where ") & " file1_10.DESCA LIKE ('%" & xDesca.Text & "%')   "
    
    cStrAll = cStrAll & " GROUP BY file1_10.item , file1_10.desca , FILE1_10.cost0  , FILE1_10.PRICE  , myitemcost.costitem order by file1_10.item "
    data4.RecordSource = cStrAll
    data4.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 23
    .RowHeight(0) = 1000
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = " ﬂ·›… ‘—ﬂ…"
    .TextMatrix(0, 3) = "”⁄— Ã„·…"
    .TextMatrix(0, 4) = " ﬂ·›… „Õ·"
    
    .TextMatrix(0, 5) = "—’Ìœ √Ê·"
    .TextMatrix(0, 6) = " ﬂ·›… —’Ìœ √Ê·"
    .TextMatrix(0, 7) = "„‘ —Ì«  ⁄œœ"
    .TextMatrix(0, 8) = "„— Ã⁄«  ⁄œœ"
    .TextMatrix(0, 9) = "’«›Ï „‘ —Ì« "
    .TextMatrix(0, 10) = "ﬁÌ„… „‘ —Ì« "
    .TextMatrix(0, 11) = "„‘ —Ì«  »ﬁÌ„…  ﬂ·›… «·‘—ﬂ…"
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = "’«›Ï ﬁÌ„… „»Ì⁄«  ›⁄·Ì…"
    
    .TextMatrix(0, 14) = " ﬂ·›… «·„»Ì⁄«  ··„Õ·"
    .TextMatrix(0, 15) = " ﬂ·›… «·„»Ì⁄«  ··‘—ﬂ…"
    
    .TextMatrix(0, 16) = "—»Õ «·„Õ· „‰ «·„»⁄Ì« "
    .TextMatrix(0, 17) = "—»Õ «·‘—ﬂ… „‰ «·„»Ì⁄« "
    
    
    .TextMatrix(0, 18) = "‰”»… «·»Ì⁄"
    
    .TextMatrix(0, 19) = "—’Ìœ «·„Õ·"
    .TextMatrix(0, 20) = "Õœ «·ÿ·»"
    .TextMatrix(0, 21) = "—’Ìœ «·‘—ﬂ…"
    
    .TextMatrix(0, 22) = "ÿ»«⁄… ·≈⁄«œ… «·ÿ·»"
    
    .ColFormat(2) = "#0.00"
    .ColFormat(3) = "#0.00"
    .ColFormat(4) = "#0.00"
    .ColFormat(5) = "#0"
    .ColFormat(6) = "#0"
    .ColFormat(7) = "#0"
    .ColFormat(8) = "#0"
    .ColFormat(9) = "#0"
    .ColFormat(10) = "#0.00"
    .ColFormat(11) = "#0.00"
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColFormat(14) = "#0.00"
    .ColFormat(15) = "#0.00"
    .ColFormat(16) = "#0.00"
    .ColFormat(17) = "#0.00"
    .ColFormat(18) = "#0.00"
    
    .ColFormat(19) = "#0"
    .ColFormat(20) = "#0"
    .ColFormat(21) = "#0"
    
    .ColHidden(2) = Not bopt1
    
    .ColHidden(4) = Not bopt1
    .ColHidden(6) = Not bopt1
    
    .ColHidden(10) = Not bopt1
    .ColHidden(11) = Not bopt1
    .ColHidden(12) = Not bopt1
    .ColHidden(13) = Not bopt1
    .ColHidden(14) = Not bopt1
    .ColHidden(15) = Not bopt1
    .ColHidden(16) = Not bopt1
    .ColHidden(17) = Not bopt1
    
    
    .AddItem "", 1
    .TextMatrix(1, 1) = "Œ’„ »Ì⁄ "
    .TextMatrix(1, 13) = Format(Val(GetDescaSHOP("select sum(discount) from file6_20h where date >= " & DateSq2(xdate1.Text) & " AND DATE <= " & DateSq2(XDATE2.Text))) * -1, "#0.00")
    .ColWidth(0) = 1800
    .ColWidth(1) = 2500
    .ColWidth(2) = 900
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .ColWidth(6) = 900
    .ColWidth(7) = 900
    .ColWidth(8) = 900
    .ColWidth(9) = 900
    
    .ColWidth(10) = 900
    .ColWidth(11) = 900
    .ColWidth(12) = 900
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 900
    .ColWidth(17) = 900
    .ColWidth(18) = 900
    .ColWidth(19) = 900
    .ColWidth(20) = 900
    .ColWidth(21) = 900
    .ColWidth(22) = 900
    .ColDataType(22) = flexDTBoolean
    
    .ColDataType(20) = flexDTBoolean
    For I = 2 To .Cols - 2
        .ColDataType(I) = flexDTDouble
    Next I

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    xcount.Caption = "⁄œœ «·«’‰«› " & Format(.Rows - 1, "#0")
    For I = 2 To .Rows - 1
        If Val(.TextMatrix(I, 2)) = 0 Then
            If xcost.Value Then .TextMatrix(I, 2) = .TextMatrix(I, 4)
            .Cell(flexcpBackColor, I, 2) = vbGreen
        End If
        
        .TextMatrix(I, 11) = Format(Val(.TextMatrix(I, 9)) * Val(.TextMatrix(I, 2)), "#0.00")
        .TextMatrix(I, 13) = Format(.TextMatrix(I, 13), "#0.00")
        .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 12)) * Val(.TextMatrix(I, 4)), "#0.00")
        .TextMatrix(I, 15) = Format(Val(.TextMatrix(I, 12)) * Val(.TextMatrix(I, 2)), "#0.00")
    
        .TextMatrix(I, 16) = Format(Val(.TextMatrix(I, 13)) - Val(.TextMatrix(I, 14)), "#0.00")
        .TextMatrix(I, 17) = Format(Val(.TextMatrix(I, 14)) - Val(.TextMatrix(I, 15)), "#0.00")
        If (Val(.TextMatrix(I, 5)) + Val(.TextMatrix(I, 9))) <> 0 Then .TextMatrix(I, 18) = Format(Val(.TextMatrix(I, 12)) / (Val(.TextMatrix(I, 5)) + Val(.TextMatrix(I, 9))) * 100, "#0.00")
        .TextMatrix(I, 20) = Format(Abs(Val(GetDescaSHOP("select f_q from f_qty where item = " & MyParn(.TextMatrix(I, 0))) & "") / 2), "#0")
        
        
    Next I
    
    .SubtotalPosition = flexSTAbove
    For I = 5 To 17
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    End With
End Sub
Private Function GetDescaSHOP(pString) As String
Dim loctable As New ADODB.Recordset
'locTable.Open pString, ConShop, adOpenStatic, adLockReadOnly, adCmdText
'If Not (locTable.BOF And locTable.EOF) Then GetDescaSHOP = locTable(0) & ""
'locTable.Close
'Set locTable = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
MyBalItem.Close
ConShop.Close
Set ConShop = Nothing
Set MyBalItem = Nothing
closeCon con
End Sub

Private Sub Grid1_DblClick()
    With grid1
        If .Col = 21 Then
            .TextMatrix(.Row, 21) = Format(Val(GetDesca("select sum([in] - [out]) as bal from file1_11 where item = " & MyParn(.TextMatrix(.Row, 0)))), "#0")
        End If
    End With
End Sub
Private Sub Grid1_EnterCell()
With grid1
    If .Col = .Cols - 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub xprint_Click()
    For I = 2 To grid1.Rows - 1
        grid1.TextMatrix(I, grid1.Cols - 1) = xprint.Value
    Next I
End Sub
Sub DoPrintItemShop()
Dim aHeader(1), lAdd As Boolean
Dim temptable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
Dim MyItemTable As New ADODB.Recordset
Dim cStr1 As String
    cStr1 = "SELECT Sum(Val(FILE1_11.[IN] & '')-Val(FILE1_11.[out] & '')) AS Balance,FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.[SECTION],FILE1_50.DESCA,FILE1_50G.DESCA,FILE1_10SC.DESCA " & _
          "FROM (((FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE) LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE WHERE TRUE "
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND [file1_10.[GROUP]]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStr1 = cStr1 & " AND file1_50.[GROUP]   = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND [Section] = " & xSection.BoundText
    If xDesca.Text <> "" Then cStr1 = cStr1 & " AND file1_10.DESCA LIKE ('%" & xDesca.Text & "%')   "
    cStr1 = cStr1 & " GROUP BY FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.[SECTION],FILE1_50.DESCA,FILE1_50G.DESCA,FILE1_10SC.DESCA  order by file1_10.item "

temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
MyItemTable.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText

aHeader(1) = "»Ì«‰ »«’‰«› „Õ· «·”·«„ ﬂ·«” ·Â« —’Ìœ »«·‘—ﬂ… Ê ·Ì” ·Â« —’Ìœ »«·„Õ· "
If (MyItemTable.EOF And MyItemTable.BOF) Then Exit Sub

With MyItemTable
    Do While Not .EOF
        If !BALANCE <> 0 Then
            lAdd = False
            nRow = grid1.FindRow(!Item, 2, 0)
            If nRow = -1 Then
                lAdd = True
            Else
                If Val(grid1.TextMatrix(nRow, 19)) <= 0 Then lAdd = True
            End If
            If lAdd Then
                temptable.AddNew
                temptable!val10 = MyItemTable!Section
                temptable!str6 = MyItemTable![file1_10SC.desca]
                temptable!val11 = MyItemTable![FILE1_50.GROUP]
                temptable!str7 = MyItemTable![file1_50G.DESCA]
                temptable!val12 = MyItemTable![FILE1_10.GROUP]
                temptable!str8 = MyItemTable![file1_50.desca]
                temptable!str1 = MyItemTable!Item
                temptable!str2 = MyItemTable![FILE1_10.DESCA]
                temptable!val2 = MyItemTable!BALANCE
                If nRow > 0 Then temptable!val9 = 1
                
                temptable!str21 = TurnValue(retHeader(aHeader, 0, 1))
                temptable.Update
            End If
        End If
        .MoveNext
    Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\Item1_SHOP.rpt"
End If
contemp.BeginTrans
contemp.CommitTrans

main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1

temptable.Close

Set temptable = Nothing
Set sourcetable = Nothing

End Sub
Private Function DateSq2(ByVal X As Variant) As Variant
If Not IsDate(X) Then
    DateSq2 = X
    Exit Function
End If
X = DateValue(Format(X, "YYYY-MM-DD"))
DateSq2 = "#" & Month(X) & "/" & Day(X) & "/" & Year(X) & "#"
End Function
Private Function myiif(cCondition, cField)
If cCondition = "" Then
    myiif = "Sum(" & cField & ")"
Else
    myiif = "Sum(iif(" & cCondition & "," & _
         cField & "," & "0" & "))"
End If
End Function
Private Function DateSq(ByVal X As Variant) As Variant
If Not IsDate(X) Then
    DateSq = X
    Exit Function
End If
X = DateValue(Format(X, "YYYY-MM-DD"))
DateSq = "#" & Month(X) & "/" & Day(X) & "/" & Year(X) & "#"
End Function

