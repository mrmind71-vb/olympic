VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form impcostfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬂ·›… «” Ì—«œÌ…"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
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
   ScaleHeight     =   10515
   ScaleWidth      =   15195
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   0
      Width           =   5145
      Begin VB.CommandButton CMDTRANS 
         Caption         =   " ÕÊÌ· ··„‘ —Ì« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   135
         Width           =   1770
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«· ”⁄Ì—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1890
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   135
         Width           =   1590
      End
      Begin VB.CommandButton cmdCalcCost 
         Caption         =   "⁄—÷ «· ﬂ·›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3510
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   135
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmditem 
      Caption         =   " ⁄œÌ· ’‰›"
      Height          =   465
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   9090
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   -3555
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2430
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   150
         Width           =   3510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9045
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄…"
         Height          =   420
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         Height          =   420
         Left            =   3870
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         Height          =   420
         Left            =   2655
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
         Height          =   420
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   1845
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   585
      Width           =   13335
      Begin VB.TextBox xDateTransPur 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   900
         Width           =   1290
      End
      Begin VB.TextBox xDateTrans 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   540
         Width           =   1290
      End
      Begin VB.TextBox xcurRate 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox xFactName 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1620
         Width           =   4395
      End
      Begin VB.TextBox xPolicy 
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
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1260
         Width           =   2445
      End
      Begin VB.TextBox xBankName 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1260
         Width           =   4395
      End
      Begin VB.CheckBox xTrans 
         Alignment       =   1  'Right Justify
         Caption         =   " —ÕÌ· «·Ì «·„Œ“‰"
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
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox xCredit 
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
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   900
         Width           =   2445
      End
      Begin VB.TextBox xVessel 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   900
         Width           =   4395
      End
      Begin VB.TextBox xCode 
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
         Height          =   315
         Left            =   11025
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1065
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
         Height          =   315
         Left            =   9855
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2235
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
         Left            =   4050
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   4050
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcurrency 
         Height          =   315
         Left            =   4050
         TabIndex        =   43
         Top             =   1620
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ÕÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ÕÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1305
         Width           =   45
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„⁄«„· «· ÕÊÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„·… :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„’‰⁄ :"
         Height          =   195
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1665
         Width           =   960
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»Ê·Ì’… :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "«”„ «·»‰ﬂ :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·«⁄ „«œ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label7 
         Caption         =   "≈”„ «·„—ﬂ» :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   990
         Width           =   1020
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   3285
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   255
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
         Height          =   195
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1530
      Width           =   1635
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   435
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         Height          =   435
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2070
      Top             =   -90
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
      Height          =   5895
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2610
      Width           =   14970
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5685
         Left            =   90
         TabIndex        =   62
         Top             =   135
         Width           =   14820
         _cx             =   26141
         _cy             =   10028
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
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
   Begin VB.Frame Frame7 
      Height          =   1380
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8505
      Width           =   11010
      Begin VB.CommandButton cmdCharge 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   900
         Width           =   330
      End
      Begin VB.TextBox xDiscount 
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
         Left            =   7650
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label xTotalNoCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "»«·⁄„·… «·„Õ·Ì… :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "»⁄œ «·Œ’„ :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
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
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label xTotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label xTotalFrgn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label xCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì «· ﬂ·›… :"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "„’«—Ì› :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "«·Œ’„ :"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   690
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì «·√’‰«› :"
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
         Index           =   0
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   630
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -2115
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
   Begin VB.Frame Frame10 
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8505
      Width           =   3930
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "»ÕÀ"
         Top             =   135
         Width           =   3750
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
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame9 
      Height          =   570
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   9000
      Width           =   1920
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -1485
      Top             =   495
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   180
      TabIndex        =   63
      Top             =   9585
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
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
End
Attribute VB_Name = "impcostfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim CardTable As ADODB.Recordset
Dim GRDTABLE2 As New ADODB.Recordset
Dim tBalance  As ADODB.Recordset
Dim cFile As String, cFileClient, cMoveName, cFileMove, cItemmove As String, cClientmove, cFieldItem, cFieldClient
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function MyReplace() As Boolean
If Not MYVALID Then Exit Function
Dim nTry As Integer, nAffected As Integer, aInsert(14, 1), aGrid(12, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xdate.Text)

aInsert(2, 0) = "code"
aInsert(2, 1) = addstring(xcode.Text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "Credit"
aInsert(4, 1) = addstring(xCredit.Text)

aInsert(5, 0) = "Vessel"
aInsert(5, 1) = addstring(xVessel.Text)

aInsert(6, 0) = "BankName"
aInsert(6, 1) = addstring(xBankName.Text)

aInsert(7, 0) = "Policy"
aInsert(7, 1) = addstring(xPolicy.Text)

aInsert(8, 0) = "FactName"
aInsert(8, 1) = addstring(xFactName.Text)

aInsert(9, 0) = "[Currency]"
aInsert(9, 1) = addstring(xcurrency.BoundText)

aInsert(10, 0) = "CurRate"
aInsert(10, 1) = addvalue(xcurRate.Text)

aInsert(11, 0) = "Trans"
aInsert(11, 1) = IIf(xTrans.Value = 0, "False", "True")

aInsert(12, 0) = "Discount"
aInsert(12, 1) = addvalue(xDiscount.Text)

aInsert(13, 0) = "Charge"
aInsert(13, 1) = addvalue(xCharge.Caption)

aInsert(14, 0) = "Datetranspur"
aInsert(14, 1) = addDate(xDateTransPur.Text)

On Error Resume Next
CON.BeginTrans
For nTry = 1 To 1
    If xDoc_No.Enabled Then
        CON.Execute CreateInsert(aInsert, "File7_60H")
    Else
        CON.Execute CreateUpdate(aInsert, "FILE7_60H", " WHERE FILE7_60H.DOC_NO = " & MyParn(xDoc_No.Text), 0)
   End If
   If Err.Number = 0 Then
       ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
        prog1.Value = 0
        prog1.Visible = True
        With Grid1
            For I = 1 To .Rows - 2
                prog1.Value = Round(I / (Grid1.Rows - 2), 2) * 100
                aGrid(0, 0) = "Doc_no": aGrid(0, 1) = addstring(xDoc_No.Text)
                aGrid(1, 0) = "item": aGrid(1, 1) = addstring(Grid1.TextMatrix(I, 1))
                aGrid(2, 0) = "Quant": aGrid(2, 1) = Val(Grid1.TextMatrix(I, 3))
                aGrid(3, 0) = "Price": aGrid(3, 1) = Val(Grid1.TextMatrix(I, 4))
                aGrid(4, 0) = "Discount": aGrid(4, 1) = Val(Grid1.TextMatrix(I, 5))
                aGrid(5, 0) = "TotalFrgn": aGrid(5, 1) = Val(Grid1.TextMatrix(I, 6))
                aGrid(6, 0) = "Total": aGrid(6, 1) = Val(Grid1.TextMatrix(I, 7))
                aGrid(7, 0) = "Cost": aGrid(7, 1) = Val(Grid1.TextMatrix(I, 8))
                aGrid(8, 0) = "Rate1": aGrid(8, 1) = Val(Grid1.TextMatrix(I, 9))
                aGrid(9, 0) = "Price1": aGrid(9, 1) = Val(Grid1.TextMatrix(I, 10))
                aGrid(10, 0) = "RATE2": aGrid(10, 1) = Val(Grid1.TextMatrix(I, 11))
                aGrid(11, 0) = "PRICE2": aGrid(11, 1) = Val(Grid1.TextMatrix(I, 12))
                aGrid(12, 0) = "ROW": aGrid(12, 1) = I
                If Grid1.TextMatrix(I, Grid1.Cols - 1) = "" Then
                    CON.Execute CreateInsert(aGrid, "FILE7_60")
                Else
                    CON.Execute CreateUpdate(aGrid, "FILE7_60", " where SR = " & Grid1.TextMatrix(I, Grid1.Cols - 1), -1)
                End If
            Next
            prog1.Visible = False
        End With
    End If
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 And nTry < 1 Then
        Err.Clear
        xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1, 15)
        aInsert(0, 1) = addstring(xDoc_No.Text)
    Else
        GoTo MyError
    End If
Next
CON.CommitTrans
MyReplace = True
Exit Function
MyError:
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = Grid1.Name Then
    nFound = Grid1.FindRow(Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    Grid1.TextMatrix(Grid1.Row, 1) = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Grid1.TextMatrix(Grid1.Row, 2) = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 1)
    GrdDesc Grid1.Row
    
    If Grid1.Row = Grid1.Rows - 1 Then
        Grid1.TextMatrix(Grid1.Rows - 1, 3) = 1
        Grid1.AddItem ""
        Grid1.Select Grid1.Rows - 1, 1
        MakeSerial
    ElseIf Grid1.Row = Grid1.Rows - 2 Then
        Grid1.TextMatrix(Grid1.Rows - 2, 3) = 1
        Grid1.Select Grid1.Rows - 1, 1
    End If
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
    CalcTotals
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    Unload Search3
    MyLoad
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Unload Search3
End If
Exit Sub
MyError:
Unload Search
End Sub
Private Sub cmdaddGroup_Click()
ReDim aPublic(0)
Set aPublic(0) = Me
additemfrm.Show 1
End Sub

Private Sub cmdCopy_Click()
If Not MYVALID Then Exit Sub
MyReplace
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
MyLoad
copyPurFrm.Show 1
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cmdCalcCost_Click()
'If Not MYVALID Then Exit Sub
'CalcCost
'If MyReplace <> 0 Then
'    MsgBox "·„ Ì „ Õ”«» «· ﬂ·›… ·ÊÃÊœ „‘ﬂ·… ›Ï «·Õ›Ÿ"
'    Exit Sub
'End If
'MsgBox "ÌÃ» Õ›Ÿ «·»Ì«‰«  «–« ÕœÀ  ⁄œÌ· ·÷„«‰ œﬁ… Õ”«»… «· ﬂ·›… «·«” Ì—«œÌ…"
'CalcCost
'CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
mySave True
doprint
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo MyError
    CON.BeginTrans
   ' Õ–› «·„” ‰œ
    CON.Execute "Delete * From FILE7_60 where Doc_No = " & MyParn(xDoc_No.Text)
    CON.Execute "Delete * From FILE7_60H where Doc_No = " & MyParn(xDoc_No.Text)
    CON.CommitTrans
    CardTable.Requery
    If CardTable.BOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    End If
End If
Exit Sub
MyError:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub

Private Sub cmdLookup2_Click()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Format([DATE],'yyyy/mm/dd'), FILE4_10.Desca,Credit,Vessel" & _
                  " FROM  (FILE7_60H left JOIN FILE4_10 ON FILE7_60H.CODE = FILE4_10.CODE )"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ-«”„ «·”›Ì‰…"
listarray(0, 1) = "(Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR %%VESSEL%% OR" & _
                  " iif(isDate('cFilter'),Format(Date,'dd-mm-yy') = Format('cFilter','dd-mm-yy'),false))"

listarray(1, 0) = "—ﬁ„ «·«⁄ „«œ"
listarray(1, 1) = "(Credit Like '%cFilter%')"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ «·«⁄ „«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«”„ «·„—ﬂ»"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "≈Ã„«·Ì «·ﬂ„Ì…"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
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
myDefine
If CardTable.EOF And CardTable.BOF Then
    xDoc_No.Text = RetZero("1", 15)
Else
    CardTable.MoveLast
    xDoc_No.Text = RetZero(IncRec(Trim(CardTable!doc_no)), 15)
End If
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
mySave
End Sub
Private Sub CmdUndo_Click()
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then
    CardTable.MoveLast
    MyLoad
Else
    MyLoad
End If
End Sub
Private Sub Cmditem_Click()
Dim bEditLocal As Boolean
bEditLocal = bEdit: bEdit = True
itemsfrm.Show 1
bEdit = bEditLocal
End Sub

Private Sub cmdCharge_Click()
impchargefrm.Show 1
xCharge.Caption = GetDesca("Select sum([value]) from fiLE7_60CH WHERE DOC_NO = " & MyParn(xDoc_No.Text))
CalcTotals
End Sub
Private Sub cmduntrans_Click()
If Not IsNull(CardTable!docPur) Then
    If myunTrans(CardTable!docPur) Then
        CardTable.Requery
        CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        MyLoad
        MsgBox " „ «·€«¡ «· —ÕÌ· »‰Ã«Õ"
    End If
End If
End Sub

Private Sub Command1_Click()
impcostpricefrm.Show 1
MyLoad
End Sub

Private Sub Command2_Click()
Dim cHeader1 As String, cHeader2 As String
contemp.Execute "delete * from print1"
printOpt.Show 1
If vPublic = 3 Then
    For I = 1 To Grid1.Rows - 2
        contemp.Execute "INSERT INTO PRINT1(ITEM,QUANT,ITEMDESCA,[GROUP],GROUPDESCA,WIDTH1,WIDTH2,MAINGROUP)" & _
                        "VALUES(" & _
                        addstring(Grid1.TextMatrix(I, 1)) & "," & _
                        addvalue(Grid1.TextMatrix(I, 3)) & "," & _
                        addstring(Grid1.TextMatrix(I, 2)) & "," & _
                        addstring(retitem(Grid1.TextMatrix(I, 1), "GROUP")) & "," & _
                        addstring(GetDesca("SELECT DESCA FROM FILE1_50 WHERE FILE1_50.CODE = " & MyParn(retitem(Grid1.TextMatrix(I, 1), "GROUP")))) & "," & _
                        addstring(retitem(Grid1.TextMatrix(I, 1), "WIDTH1") & " x " & retitem(Grid1.TextMatrix(I, 1), "WIDTH1")) & "," & _
                        addstring(IIf(IsNull(retitem(Grid1.TextMatrix(I, 1), "Length")), "No Length", retitem(Grid1.TextMatrix(I, 1), "Length"))) & "," & _
                        addstring(GetDesca("SELECT [GROUP] FROM FILE1_50 WHERE FILE1_50.CODE = " & MyParn(retitem(Grid1.TextMatrix(I, 1), "GROUP")))) & ")"
    Next
    cHeader1 = "[ " & "›« Ê—… «” Ì—«œÌ… —ﬁ„ : " & xDoc_No.Text & " ]" & Space(3) & _
               "[ " & xStore.Text & " ]" & Space(3) & _
               "[ " & " » «—ÌŒ : " & xdate.Text & " ]" & Space(3)
    cHeader2 = turnFound("[ ", xVessel.Text) & xVessel.Text & turnFound(xVessel.Text) & "«·„—ﬂ» : " & turnFound(xVessel.Text, " ]") & turnFound(xVessel.Text, Space(3)) & "[ " & xcodedesca.Caption & "«·„Ê—œ " & " ]"
    printTablefrm3.nDigitAdd = 2
    printTablefrm3.doprint cHeader1, cHeader2, , True, False
    printTablefrm3.Show 1
Else
    For I = 1 To Grid1.Rows - 2
        contemp.Execute "INSERT INTO PRINT1(ITEM,QUANT,ITEMDESCA,[GROUP],GROUPDESCA,WIDTH1,WIDTH2,MAINGROUP)" & _
                        "VALUES(" & _
                        addstring(Grid1.TextMatrix(I, 1)) & "," & _
                        addvalue(Grid1.TextMatrix(I, 3)) & "," & _
                        addstring(Grid1.TextMatrix(I, 2)) & "," & _
                        addstring(retitem(Grid1.TextMatrix(I, 1), "GROUP")) & "," & _
                        addstring(GetDesca("SELECT DESCA FROM FILE1_50 WHERE FILE1_50.CODE = " & MyParn(retitem(Grid1.TextMatrix(I, 1), "GROUP")))) & "," & _
                        addstring(retitem(Grid1.TextMatrix(I, 1), "WIDTH1")) & "," & _
                        addstring(retitem(Grid1.TextMatrix(I, 1), "WIDTH2")) & "," & _
                        addstring(GetDesca("SELECT [GROUP] FROM FILE1_50 WHERE FILE1_50.CODE = " & MyParn(retitem(Grid1.TextMatrix(I, 1), "GROUP")))) & ")"
    Next
    cHeader1 = "[ " & "›« Ê—… «” Ì—«œÌ… —ﬁ„ : " & xDoc_No.Text & " ]" & Space(3) & _
               "[ " & xStore.Text & " ]" & Space(3) & _
               "[ " & " » «—ÌŒ : " & xdate.Text & " ]" & Space(3)
    cHeader2 = "[ " & xVessel.Text & " : «·„—ﬂ»" & " ]" & Space(3) & "[ " & xcodedesca.Caption & " : «·„Ê—œ" & " ]"
               
    
    'cHeader1 = " „” ‰œ Ã—œ —ﬁ„ : " & xDoc_No.Text: cheader2 = xStore.Text: cheader3 = " » «—ÌŒ : " & xDate.Text
    If vPublic = 0 Then
        printTablefrm.nDigitAdd = 2
        printTablefrm.doprint cHeader1, cHeader2, , True, False
        printTablefrm.Show 1
    ElseIf vPublic = 1 Then
        printTablefrm.nDigitAdd = 2
        printTablefrm2.doprint cHeader1, cHeader2, , True, False
        printTablefrm2.Show 1
    ElseIf vPublic = 2 Then
        doprint2
    End If
End If
End Sub
Private Sub CMDTRANS_Click()
If Not myValidTrans Then Exit Sub
cString = InputBox(" «—ÌŒ «· —ÕÌ· ", " —ÕÌ·  ﬂ·›… «” Ì—«œÌ…", xdate.Text)
If Not IsDate(cString) Then
    MsgBox IIf(Trim(cString) = "", " „  Ã«Â· «· —ÕÌ·", "«· «—ÌŒ €Ì— ’ÕÌÕ")
    Exit Sub
End If
xDateTransPur.Text = cString
If MsgBox("Â·  Êœ «·Õ›Ÿ ﬁ»· «· —ÕÌ·", vbYesNo + vbDefaultButton1) = vbYes Then
    If mySave Then MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ Ê”Ì „ «· —ÕÌ· «·¬‰"
End If
If myTrans = 0 Then
    MsgBox " „ «· —ÕÌ· »‰Ã«Õ"
    CardTable.Requery
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    MyLoad
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
cFile = "File7_60"
cFileHeader = "File7_60H"
cFileClient = "File4_10"
With Grid1
    .Cols = 13
    .RowHeight(0) = 700
    .Editable = flexEDKbd
    .WordWrap = True
End With
'cmditem.Enabled = RetSec("tmItem")
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM FILE7_60H  ORDER BY DOC_NO", CON, adOpenKeyset, adLockReadOnly, adCmdText

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

data2.ConnectionString = CON.ConnectionString
data2.RecordSource = "FILE0_60"
Set xcurrency.RowSource = data2
xcurrency.ListField = "Desca"
xcurrency.BoundColumn = "Code"

Set Grid1.DataSource = data3
data3.ConnectionString = CON.ConnectionString
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    myDefine
    FixGrd
    xDoc_No.Text = RetZero("1", 15)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Search3
GRDTABLE.Close
Set GRDTABLE = Nothing
GRDTABLE2.Close
Set GRDTABLE2 = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Grid1.Col = 1 Then
    GrdDesc Row
End If
CalcTotals
End Sub
Private Sub Grid1_EnterCell()
If (Grid1.Col = 0 Or Grid1.Col = 2 Or Grid1.Col = 6 Or Grid1.Col = 7 Or Grid1.Col = 8 Or Grid1.Col = 9) Then
    Grid1.Editable = flexEDNone
Else
    Grid1.Editable = flexEDKbd
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And Grid1.Row <> Grid1.Rows - 1 Then
    Grid1.AddItem "", Grid1.Row
    MakeSerial Grid1.Row - 1
End If
If KeyCode = 112 And Grid1.Col = 1 Then
    ItemsLookupAll Me, Search3
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Grid1.Row = Grid1.Rows - 1 Then
    Grid1.AddItem ""
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
xcode.Text = ""
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "Select code ,DescA From FILE4_10"
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
End If
End Sub

Private Sub xCode_LostFocus()
xcodedesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xcode.Text)) & ""
End Sub

Private Sub xDateTrans_Change()
CMDTRANS.Enabled = IsDate(xDateTrans.Text)
End Sub

Private Sub xDiscount_LostFocus()
CalcTotals
End Sub
Private Function MYVALID() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xdate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xcode.Text = "" Then
    MsgBox "·„ Ì „ «œŒ«· ≈”„ «·⁄„Ì·"
    Exit Function
End If

If (Not IsDate(xDateTrans.Text)) And xTrans.Value <> 0 Then
    MsgBox " «—ÌŒ «· —ÕÌ· €Ì— ”·Ì„ «Ê €Ì— „ÊÃÊœ"
    Exit Function
End If
With Grid1
For I = 1 To Grid1.Rows - 2
    If Not validRow(I) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
If xTrans.Value <> 0 Then
    cString = Trim(GetDesca("Select Doc_no from file7_20H where docimp = " & MyParn(xDoc_No.Text)) & "")
    If cString <> "" Then
        MsgBox "·‰ Ì „ «·Õ›Ÿ «·—Ã«¡ «€·«ﬁ «Œ Ì«— «· —ÕÌ· ÕÌÀ «‰Â  „  —ÕÌ· «·„” ‰œ „‰ ﬁ»·" & vbCrLf & _
               " „” ‰œ „‘ —Ì«  —ﬁ„ : " & cString
        Exit Function
    End If
End If
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xdate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!Store
xcode.Text = CardTable!CODE
xTrans.Value = IIf(CardTable!TRANS, 1, 0)
xCredit.Text = CardTable!CREDIT & ""
xVessel.Text = CardTable!Vessel & ""
xBankName.Text = CardTable!BankName & ""
xPolicy.Text = CardTable!Policy & ""
xcurrency.BoundText = CardTable!Currency & ""
xDateTransPur.Text = Format(CardTable!DateTransPur, "dd-mm-yyyy")
xcurRate.Text = CardTable!curRate & ""
xFactName.Text = CardTable!FactName & ""
xcodedesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xcode.Text)) & ""
xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = Format(CardTable!Discount & "", "Fixed")
xCharge.Caption = GetDesca("Select sum([value]) from fiLE7_60CH WHERE DOC_NO = " & MyParn(xDoc_No.Text))

cString = "SELECT FILE7_60.ROW,FILE1_10.ITEM,FILE1_10.DESCA,FILE7_60.Quant,FILE7_60.Price,FILE7_60.DISCOUNT,0 AS TOTALFRGN,FILE7_60.TOTAL,FILE7_60.COST,FILE7_60.RATE1,FILE7_60.PRICE1,FILE7_60.RATE2,FILE7_60.PRICE2,FILE7_60.SR" & _
          " FROM FILE7_60 INNER JOIN FILE1_10 ON FILE7_60.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & _
          " ORDER BY FILE7_60.ROW"
data3.RecordSource = cString
data3.Refresh
Grid1.AddItem ""
Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
Handlecontrols LoadMode
CalcTotals
FixGrd
End Sub
Private Sub myDefine()
xdate.Text = ""
xCredit.Text = ""
xVessel.Text = ""
xTrans.Value = 0
xBankName.Text = ""
xPolicy.Text = ""
xcurrency.BoundText = ""
xDateTransPur.Text = ""
xcurRate.Text = ""
xFactName.Text = ""
xStore.BoundText = ""
xcodedesca.Caption = ""
xcode.Text = ""
xDiscount.Text = ""
xTotal.Caption = ""
xTotalItem.Caption = ""
xusername.Text = ""
Grid1.Rows = 1
Grid1.AddItem ""
Grid1.TextMatrix(1, 0) = 1
FixGrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
'If Not (CardTable.EOF And CardTable.BOF) Then cmdTrans.Enabled = IsNull(CardTable!docPur) And nMode = LoadMode Else cmdTrans.Enabled = False
'If Not (CardTable.EOF And CardTable.BOF) Then cmduntrans.Enabled = (Not IsNull(CardTable!docPur)) And nMode = LoadMode And bEdit Else cmdTrans.Enabled = False
'xTrans.Enabled = Not cmduntrans.Enabled

cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = (nMode = LoadMode) And bEdit
'cmdSave.Enabled = (bEdit) And (CanEdit) Or nMode = DefineMode
'CmdDelInv.Enabled = (nMode = LoadMode And CanEdit) And bDel
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text, 15)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Grid1.Row <> Grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ", vbOKCancel + vbDefaultButton2) = vbOK Then
        If Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) <> "" Then CON.Execute "DELETE * FROM FILE7_60CH WHERE SR = " & Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)
        Grid1.RemoveItem Grid1.Row
        CalcTotals
        CalcCost
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case Grid1.Col
    Case 1
        If KeyCode = 27 Then Exit Sub
End Select
End Sub
Private Sub GrdDesc(Row)
If Grid1.Col <> 1 Then Exit Sub
For nCol = 2 To Grid1.Cols - 2
    Grid1.TextMatrix(Row, nCol) = ""
Next
If Grid1.TextMatrix(Row, 1) = "" Then Exit Sub
rdItem.Find "ITEM = " & MyParn(Grid1.TextMatrix(Row, 1)), , adSearchForward, adBookmarkFirst
If Not rdItem.EOF Then
    Grid1.TextMatrix(Row, 2) = rdItem!Desca & ""
End If
'If grid1.Row > 1 Then If SameGroup(grid1.TextMatrix(Row, 1), grid1.TextMatrix(Row - 1, 1)) Then grid1.TextMatrix(Row, 4) = grid1.TextMatrix(Row - 1, 4)
End Sub
Private Function CalcTotals()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalFrgn As Double, nTotalNoCharge
With Grid1
For I = 1 To Grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    Grid1.TextMatrix(I, 6) = Round(Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount, 2)
    nTotalitem = nTotalitem + Val(Grid1.TextMatrix(I, 6))
    nTotalQuant = nTotalQuant + Val(Grid1.TextMatrix(I, 3))
Next
nTotalFrgn = nTotalitem - Val(xDiscount.Text)
nTotalNoCharge = nTotalFrgn * Val(xcurRate.Text)
xTotalItem.Caption = Format(nTotalitem, "Fixed")
xTotalFrgn.Caption = Format(nTotalFrgn, "Fixed")
xTotalNoCharge.Caption = Format(nTotalNoCharge, "Fixed")
xTotal.Caption = (nTotalFrgn * Val(xcurRate.Text)) + Val(xCharge.Caption)
xTotalQuant.Caption = Format(nTotalQuant, "#0.0000")
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 4)
Dim GrdArray(8, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Format([DATE],'yyyy/mm/dd'), " & cFileClient & ".Desca,Credit,Vessel," & cFileHeader & ".TotalQuant,iif(isNull(docPur),'€Ì— „—Õ·','„—Õ·') , iif(trans,FILE0_40.DESCA,' ') " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE ) LEFT JOIN FILE0_40 ON " & cFileHeader & ".store = file0_40.code"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ-«”„ «·”›Ì‰…"
listarray(0, 1) = "(Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR %%VESSEL%% OR" & _
                  " iif(isDate('cFilter'),Format(Date,'dd-mm-yy') = Format('cFilter','dd-mm-yy'),false))"

listarray(1, 0) = "—ﬁ„ «·«⁄ „«œ"
listarray(1, 1) = "(Credit Like '%cFilter%')"

listarray(2, 0) = "«·„Œ“‰"
listarray(2, 1) = "(STORE = 'cFilter')"
listarray(2, 2) = "FILE0_40"
listarray(2, 3) = "CODE"
listarray(2, 4) = "DESCA"

'listarray(2, 0) = "«·ﬂ„Ì…"
'listarray(2, 1) = "(Fix(Totalquant) Like Fix(val(cFilter)))"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ «·«⁄ „«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«”„ «·„—ﬂ»"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "≈Ã„«·Ì «·ﬂ„Ì…"
GrdArray(6, 1) = 1500

GrdArray(7, 0) = "«· —ÕÌ·"
GrdArray(7, 1) = 1100

GrdArray(8, 0) = "„Œ“‰"
GrdArray(8, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To Grid1.Rows - 2
    If I <> nRow And Trim(Grid1.TextMatrix(I, nCol)) <> "" Then
        If Trim(Grid1.TextMatrix(I, nCol)) = Trim(Grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For I = 1 To Grid1.Rows - 2
    nRow = FoundOtherRow(I, 1)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & Grid1.TextMatrix(Grid1.Row, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function
Private Sub xfilter_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd Grid1, xfilter(Index), 2
End If
End Sub

Private Sub xRate_LostFocus()
CalcTotals
End Sub
Private Function validRow(nRow) As Boolean
If nRow > 0 Then
    If Trim(Grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
    If Trim(Grid1.TextMatrix(nRow, 2)) = "" Then Exit Function
   ' If Val(grid1.TextMatrix(nRow, 5)) = 0 Then Exit Function
End If
validRow = True
End Function
Sub additemProc()
Grid1.RemoveItem Grid1.Rows - 1
With additemfrm.Grid1
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(I, 4)) <> 0 Then
            Grid1.AddItem ""
            Grid1.TextMatrix(Grid1.Rows - 1, 1) = .TextMatrix(I, 0)
            Grid1.TextMatrix(Grid1.Rows - 1, 2) = retitem(.TextMatrix(I, 0), "desca")
            Grid1.TextMatrix(Grid1.Rows - 1, 3) = .TextMatrix(I, 4)
            Grid1.TextMatrix(Grid1.Rows - 1, 4) = .TextMatrix(I, 5)
            Grid1.TextMatrix(Grid1.Rows - 1, 7) = retitem(.TextMatrix(I, 0), "width1") & ""
            Grid1.TextMatrix(Grid1.Rows - 1, 8) = retitem(.TextMatrix(I, 0), "width2") & ""
            Grid1.TextMatrix(Grid1.Rows - 1, 9) = retitem(.TextMatrix(I, 0), "length") & ""
        End If
    Next
    Grid1.AddItem ""
    CalcTotals
End With
End Sub
Private Sub addPur()
Dim locTable As New ADODB.Recordset
locTable.Open "file7_41", CON, adOpenStatic, adLockReadOnly, adCmdTable
locTable.Find "doc_no = " & MyParn(xDocpur.Text), , adSearchForward, adBookmarkFirst

Grid1.Rows = 2
If Not locTable.EOF Then
    xdate.Text = Format(locTable!Date, "dd-mm-yyyy")
    xStore.BoundText = locTable!Store
    xcode.Text = locTable!CODE
    xcodedesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xcode.Text)) & ""
    xusername.Text = TurnValue(locTable!UserName, Null, "")
    Grid1.Rows = 1
    Dim LocGrdTable As New ADODB.Recordset
    LocGrdTable.Open "select * from file7_40 where doc_no = " & MyParn(xDocpur.Text), CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not (LocGrdTable.EOF And LocGrdTable.BOF) Then
        With Grid1
            Do
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = TurnValue(LocGrdTable!Item, Null, "")
                .TextMatrix(.Rows - 1, 1) = TurnValue(retitem(LocGrdTable!Item, "DescA"), Null, "")
                .TextMatrix(.Rows - 1, 2) = LocGrdTable!Quant & ""
                .TextMatrix(.Rows - 1, 6) = retitem(LocGrdTable!Item, "width1") & ""
                .TextMatrix(.Rows - 1, 7) = retitem(LocGrdTable!Item, "width2") & ""
                .TextMatrix(.Rows - 1, 8) = retitem(LocGrdTable!Item, "Length") & ""
                 LocGrdTable.MoveNext
            Loop Until LocGrdTable.EOF
        End With
    End If
    Grid1.AddItem ""
    LocGrdTable.Close
End If
locTable.Close
End Sub
Private Function RetItemBalance(citem, cStore, dDate) As Double
If citem = "" Then Exit Function
movetable.Seek Array(citem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> citem Or cStore <> movetable!Store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    If Not (movetable!Type = cItemmove And movetable!doc_ID = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    End If
    movetable.MoveNext
Loop
End Function
Private Sub CalcCost()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalFrgn As Double, nTotalNoCharge
With Grid1
For I = 1 To Grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    Grid1.TextMatrix(I, 6) = Round(Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount, 2)
    nTotalitem = nTotalitem + Val(Grid1.TextMatrix(I, 6))
    nTotalQuant = nTotalQuant + Val(Grid1.TextMatrix(I, 3))
Next
nTotalFrgn = nTotalitem - Val(xDiscount.Text)
nTotalNoCharge = nTotalFrgn * Val(xcurRate.Text)
nTotal = nTotalNoCharge + Val(xCharge.Caption)

If Val(nTotalFrgn) = 0 Then Exit Sub
nRate = nTotal / nTotalFrgn
For I = 1 To Grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    Grid1.TextMatrix(I, 8) = Round(Val(Grid1.TextMatrix(I, 4)) * nDiscount * nRate, 2)
    Grid1.TextMatrix(I, 7) = Round(Val(Grid1.TextMatrix(I, 8)) * Val(Grid1.TextMatrix(I, 3)), 2)
Next
End With
End Sub
Private Sub doprint()
Dim aHeader(3)
If Val(xTotalFrgn.Caption) = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ’ÕÌÕ… ·ÿ»«⁄ Â«"
    Exit Sub
End If
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset

contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.GROUP,FILE7_60.ITEM,File1_10.Desca,FILE7_60.CODE, FILE7_60.Quant, FILE7_60.COST,FILE7_60.PRICE,FILE7_60.PRICE1,FILE7_60.PRICE2  " & _
          " FROM FILE1_10 INNER JOIN FILE7_60 ON FILE1_10.ITEM = FILE7_60.ITEM" & _
          " WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text) & _
          " ORDER BY FILE7_60.ROW "
         
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

aHeader(0) = "[" & "›« Ê—… —ﬁ„ : " & xDoc_No.Text & "]"
aHeader(1) = "[" & "» «—ÌŒ : " & xdate.Text & "]"
aHeader(2) = "[" & "··„Ê—œ : " & xcodedesca.Caption & "]"
If xVessel.Text <> "" Then aHeader(3) = "[" & " «·”›Ì‰… : " & xVessel.Text & "]"
With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!Val10 = Val(xTotalFrgn.Caption)
        temptable!Val11 = Val(xcurRate.Text)
        temptable!Val12 = Val(xCharge.Caption)
        temptable!Val13 = Val(xTotal.Caption)
        temptable!VAL14 = (Val(xTotal.Caption)) / Val(xTotalFrgn.Caption)
        temptable!str1 = !Item
        temptable!str2 = !Desca
        temptable!val1 = !Quant
        temptable!val2 = !price
        temptable!val3 = Val(!Quant & "") * Val(!price & "")
        temptable!val4 = !cost
        temptable!val5 = Val(!Quant & "") * Val(!cost & "")
        temptable!Val6 = !Price1
        temptable!Val7 = Val(!Quant & "") * Val(!Price1 & "")
        temptable!Val8 = !price2
        temptable!val9 = Val(!Quant & "") * Val(!price2 & "")
        temptable!str21 = retHeader(aHeader, 0, 2)
        temptable!str22 = retHeader(aHeader, 2, 2)
        temptable.Update
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\impcost_1.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub FixGrd()
With Grid1
    .Editable = flexEDKbd
    .FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·Œ’„|" & "«·≈Ã„«·Ì »«·⁄„·…|" & "≈Ã„«·Ì «· ﬂ·›… »«·Ã‰ÌÂ|" & " ﬂ·›… «·ÊÕœ… »«·Ã‰ÌÂ|" & "‰”»… «·Ã„·…|" & "”⁄— «·Ã„·…|" & "‰”»… „” Â·ﬂ|" & "”⁄— „” Â·ﬂ|"
    .ColWidth(0) = 500
    .ColWidth(1) = 1600
    .ColWidth(2) = 3000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColHidden(Grid1.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = nBeginRow To Grid1.Rows - 1
    Grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Function myunTrans(pDoc_No) As Boolean
On Error GoTo MyError
CON.BeginTrans
CON.Execute "DELETE * FROM FILE7_21 " & _
        " WHERE FILE7_21.DOC_NO = " & MyParn(pDoc_No)

CON.Execute "DELETE * FROM FILE7_20 " & _
        " WHERE FILE7_20.DOC_NO = " & MyParn(pDoc_No)

CON.Execute "DELETE * FROM FILE1_11 " & _
           " WHERE TYPE = '2' AND DOC_ID = " & MyParn(pDoc_No)

CON.Execute "DELETE * FROM FILE4_11 WHERE DOC_ID = " & MyParn(pDoc_No) & " AND TYPE = '2'"
CON.Execute "update FILE7_60H set FILE7_60H.docpur = null ,dateTrans = null WHERE DOC_NO = " & MyParn(xDoc_No.Text), nRecord
CON.CommitTrans
myunTrans = True
Exit Function
MyError:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function myTrans() As Boolean
Dim nTry As Integer
On Error Resume Next
cDoc_No = RetZero(Newflag("FILE7_20H", "doc_no"), 6)
CON.BeginTrans
Do
    Err.Clear
    nTry = nTry + 1
    CON.Execute "INSERT INTO FILE7_20h (DOC_NO, code, store, [date], discount,DocImp) " & _
                " SELECT " & MyParn(cDoc_No) & ", FILE7_60H.code, FILE7_60H.store, " & addDate(xDateTransPur.Text) & ", FILE7_60H.discount,FILE7_60H.DOC_NO" & _
                " FROM FILE7_60H " & _
                " WHERE FILE7_60H.DOC_NO = " & MyParn(xDoc_No.Text)
    
    If Err.Number = 0 Then
        CON.Execute "INSERT INTO FILE7_20 ( DOC_NO, ITEM, Quant, PRICE, Total,[ROW]) " & _
              " SELECT " & MyParn(cDoc_No) & ", FILE7_60.ITEM, FILE7_60.Quant,FILE7_60.cost, Val([quant] & '') * Val([cost] & ''),[ROW]" & " FROM FILE7_60 " & _
              " WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text)
    End If
    If Err.Number = 0 Then Exit Do
    If Err.Number = -2147467259 And nTry < 10 Then
        cDoc_No = RetZero(Format(cDoc_No) + 1, 15)
    Else
        GoTo MyError
    End If
Loop Until Err.Number = 0
CON.CommitTrans
myTrans = True
Exit Function
MyError:
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function

Private Function CanEdit() As Boolean
If Not (CardTable.EOF And CardTable.BOF) Then If Not IsNull(CardTable!docPur) Then Exit Function
CanEdit = True
End Function
Private Function myValidTrans() As Boolean
Dim cString
cString = GetDesca("Select Doc_no from file7_20H where docimp = " & MyParn(xDoc_No.Text)) & ""
If Trim(cString) <> "" Then
    MsgBox "Â‰«ﬂ „” ‰œ „‘ —Ì«   „  —ÕÌ·… »—ﬁ„ " & cString
    Exit Function
End If
myValidTrans = True
End Function
Private Sub doprint2()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(0)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA AS ITEMDESCA,FILE1_10.UNIT, " & _
          "  FILE7_60.DOC_NO,FILE7_60H.CODE as SupCode,FILE7_60H.DATE,FILE7_60H.VESSEL, " & _
          " Sum(val(FILE7_60.Quant & '')) AS SumofQuant, " & _
          " FILE1_10.GROUP AS GroupCode, FILE1_50.DESCA AS GroupDesca,  " & _
          " FILE1_50.GROUP AS MainGroupCode, FILE1_51.DESCA as  MainGroupDesca" & _
          " FROM (((FILE7_60 INNER JOIN FILE7_60H ON FILE7_60.DOC_NO = FILE7_60H.DOC_NO )INNER JOIN FILE1_10  ON FILE7_60.ITEM = FILE1_10.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_51 ON " & _
          " FILE1_50.GROUP = FILE1_51.CODE"

cString = cString & turnFound(cString) & "FILE7_60H.doc_no = " & MyParn(xDoc_No.Text)
cString = cString & " Group by  FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA,FILE1_10.UNIT," & _
          " FILE7_60.DOC_NO,FILE7_60H.CODE,FILE7_60H.DATE,FILE7_60H.VESSEL, " & _
          " FILE1_10.GROUP, FILE1_50.DESCA,FILE1_50.GROUP, FILE1_51.DESCA  "

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!STR6 = !MainGroupDesca
        temptable!str5 = !MAINGROUPCODE
        temptable!str1 = !Item
        temptable!str2 = itemWidth(sourcetable!Item)
        temptable!str3 = !GroupCode
        temptable!str4 = !GroupDesca
        temptable!str8 = GetDesca("Select Desca from file1_13 where code = " & MyParn(!unit))
        temptable!str9 = !doc_no
        temptable!str10 = GetDesca("Select Desca from file4_10 where code = " & MyParn(!supCode))
        temptable!Str11 = !Vessel
        temptable!Date1 = !Date
        temptable!val1 = !sumOfQuant
        temptable!VAL20 = !width1
        temptable!STR20 = !width1
        temptable!str17 = TurnValue(retHeader(aHeader, 0, 4))
        temptable.Update
        .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\invdtl.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Function itemWidth(pItem) As String
itemWidth = retitem(pItem, "width1") & ""
If Not IsNull(retitem(pItem, "width2")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "width2")
If Not IsNull(retitem(pItem, "length")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "length")
End Function
Function mySave(Optional bIgMsg As Boolean = False)
If Not MYVALID Then Exit Function
CalcCost
CalcTotals
If Not MyReplace Then Exit Function
CardTable.Requery
If Not bIgMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
MyLoad
mySave = True
End Function
