VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form SalesFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9765
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
   ScaleHeight     =   9765
   ScaleWidth      =   15195
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   600
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   0
      Width           =   6000
      Begin VB.CommandButton cmdBarCode 
         Caption         =   " ÕÊÌ· ·»«—ﬂÊœ"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton cmdBarCodeopen 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄… «–‰ ’—›"
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
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… ›« Ê—…"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "≈Ã„«·Ì „»Ì⁄«  ÌÊ„Ì…"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   135
         Visible         =   0   'False
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   5490
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
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
         Left            =   4095
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
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
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
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
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·„” ‰œ"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1680
      Left            =   1755
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   585
      Width           =   13380
      Begin VB.TextBox xNotes 
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
         Left            =   6255
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1215
         Width           =   5865
      End
      Begin VB.TextBox xCode 
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
         Left            =   11025
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   1095
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
         Height          =   330
         Left            =   11025
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1095
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xman 
         Height          =   315
         Left            =   9675
         TabIndex        =   4
         Top             =   855
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   90
         TabIndex        =   49
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xbalanceitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1260
         Width           =   2445
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «·’‰› :"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1305
         Width           =   1035
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   12255
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1305
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·»«∆⁄ :"
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
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   945
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   495
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   495
         Width           =   3165
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   12210
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
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
         Left            =   12210
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   600
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   180
      TabIndex        =   40
      Top             =   9315
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame3 
      Height          =   960
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1305
      Width           =   1545
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         CausesValidation=   0   'False
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Width           =   1365
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2205
      Width           =   14970
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   6255
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   14820
         _cx             =   26141
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         RightToLeft     =   0   'False
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
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8730
      Width           =   1905
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
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Move Last"
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
         TabIndex        =   30
         TabStop         =   0   'False
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
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Width           =   435
      End
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2745
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
   Begin VB.Frame Frame7 
      Height          =   960
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8775
      Width           =   8925
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox xRateDis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   540
         Width           =   510
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6525
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   540
         Width           =   780
      End
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
         TabIndex        =   8
         Top             =   990
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label14 
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
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Width           =   690
      End
      Begin VB.Label xDisItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label10 
         Caption         =   "’«›Ì »⁄œ «·Œ’„  «·‰ﬁœÌ :"
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
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   495
         Width           =   1275
      End
      Begin VB.Label xTotalDis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
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
         TabIndex        =   39
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label lblTotalQuant 
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
         Left            =   1395
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label xTotalDisItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
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
         TabIndex        =   37
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label12 
         Caption         =   "’«›Ì «·›« Ê—… :"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   165
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   630
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "Œ’„ ‰ﬁœÌ :"
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
         Left            =   7410
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   645
         Width           =   1365
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5970
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Left            =   7410
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "÷—«∆» «·„»Ì⁄«  :"
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
         Left            =   4605
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -1260
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
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   1845
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   135
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   4095
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   90
      Top             =   360
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
   Begin MSAdodcLib.Adodc DATA4 
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
   Begin VB.Frame Frame5 
      Height          =   960
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   8775
      Width           =   4065
      Begin VB.CommandButton cmdVisa 
         Caption         =   "›Ì“«"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdCash 
         Caption         =   "‰ﬁœÌ"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox xVisa 
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
         Left            =   2115
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xCash 
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
         Left            =   2115
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label15 
         Caption         =   "¬Ã· :"
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
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   495
      End
   End
End
Attribute VB_Name = "SalesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset, ClientTable As ADODB.Recordset, cFileHeader As String, rdPaid As New ADODB.Recordset
Public bRetvalue As Boolean
Dim ItemTable As New ADODB.Recordset, bInsert As Boolean, bAddnew As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String
Dim GRDTABLE As New ADODB.Recordset
Dim SEARCH31 As New Search3, search32 As New Search3, bMarket As Boolean
Dim dLastdate As String, bEdit As Boolean
Dim cFile As String, cFileClient, cCodeDesca As String
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.price,file1_10.price2 From file1_10 left join file1_50 on file1_10.group = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "”⁄— «·Ã„·…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "«·ﬁÿ«⁄Ì"
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Function MyReplace() As Boolean
Dim nTry As Integer
On Error Resume Next
CON.BeginTrans
For nTry = 1 To 10
    If xDoc_no.Enabled Then
        CON.Execute "insert into " & cFileHeader & "(doc_no,[date],code,store,man,inv_no,Discount,Tax,cash,visa,total,Notes,username) values(" & _
                    addstring(xDoc_no.Text) & "," & _
                    DateSq(xdate.Text) & "," & _
                    addstring(xCode.Text) & "," & _
                    addstring(xStore.BoundText) & "," & _
                    addstring(xman.BoundText) & "," & _
                    Val(xDiscount.Text) & "," & _
                    Val(xTax.Text) & "," & _
                    Val(xCash.Text) & "," & _
                    Val(xVisa.Text) & "," & _
                    addvalue(xtotal.Caption) & "," & _
                    addstring(xNotes.Text) & "," & _
                    addstring(sUserName) & _
                    ")"
    Else
        CON.Execute "update " & cFileHeader & " set " & _
                    "[DATE] = " & DateSq(xdate.Text) & _
                    ",CODE = " & addstring(xCode.Text) & _
                    ",store = " & addstring(xStore.BoundText) & _
                    ",man = " & addstring(xman.BoundText) & _
                    ",DISCOUNT = " & Val(xDiscount.Text) & _
                    ",TAX = " & Val(xTax.Text) & _
                    ",CASH = " & Val(xCash.Text) & _
                    ",VISA = " & Val(xVisa.Text) & _
                    ",TOTAL = " & addvalue(xtotal.Caption) & _
                    ",Notes = " & addstring(xNotes.Text) & _
                    ",username = " & addstring(sUserName) & _
                    " where doc_no = " & MyParn(xDoc_no.Text)
    End If
    If Err.Number = 0 Then
           ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
        CON.Execute " Delete * From " & cFile & " where Doc_No = " & MyParn(xDoc_no.Text) & " and Row > " & Grid1.Rows - 2
        prog1.Value = 0
        prog1.Visible = True
        With Grid1
            For I = 1 To .Rows - 2
                prog1.Value = Round(I / (Grid1.Rows - 2), 2) * 100
                CON.Execute "Insert Into " & cFile & " (Doc_no,Item,Quant,Price,Total,Discount,row,Cost) " & _
                            " Values(" & _
                            addstring(xDoc_no.Text) & "," & _
                            addstring(.TextMatrix(I, 1)) & "," & _
                            addvalue(.TextMatrix(I, 3)) & "," & _
                            addvalue(.TextMatrix(I, 4)) & "," & _
                            Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) & "," & _
                            Val(.TextMatrix(I, 6)) & "," & _
                            I & "," & _
                            Val(itemCost(Grid1.TextMatrix(I, 1), xdate.Text)) & _
                            ")"
                If Err.Number = -2147467259 Then
                    Err.Clear
                    CON.Execute "update " & cFile & _
                                " set item = " & addstring(.TextMatrix(I, 1)) & _
                                ", Quant = " & Val(.TextMatrix(I, 3)) & _
                                ", Price = " & Val(.TextMatrix(I, 4)) & _
                                ", Total = " & Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) & _
                                ", discount = " & Val(.TextMatrix(I, 6)) & _
                                ",Cost = " & Val(itemCost(Grid1.TextMatrix(I, 1), xdate.Text)) & _
                                " where Doc_no = " & MyParn(xDoc_no.Text) & " and Row = " & I, nRecord
                End If
                If Err.Number <> 0 Then GoTo myerror
            Next
            prog1.Visible = False
        End With
    End If
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 And nTry < 10 Then
        Err.Clear
        xDoc_no.Text = RetZero(Val(xDoc_no.Text) + 1)
    End If
    If Err.Number <> 0 Then GoTo myerror
Next
CON.CommitTrans
MyReplace = True
Exit Function
myerror:
prog1.Visible = False
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = Grid1.Name Then
    nFound = Grid1.FindRow(Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    Grid1.EditText = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Grid1.TextMatrix(Grid1.Row, 1) = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Grid1.TextMatrix(Grid1.Row, 2) = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 1)
    Grid1.TextMatrix(Grid1.Row, 3) = ""
    GrdDesc Grid1.Row
    If Grid1.Row = Grid1.Rows - 1 Then
        Grid1.TextMatrix(Grid1.Rows - 1, 3) = ""
        Grid1.AddItem ""
        MakeSerial
        Grid1_AfterEdit Grid1.Row, Grid1.Col
        Grid1.Select Grid1.Rows - 1, 1
    Else
        Grid1.TextMatrix(Grid1.Row, 3) = ""
        Grid1_AfterEdit Grid1.Row, Grid1.Col
        Grid1.Select Grid1.Rows + 1, 1
    End If
    CalcTotals
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(SEARCH31.Grid1.TextMatrix(SEARCH31.Grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    SEARCH31.Hide
    MyLoad
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = search32.Grid1.TextMatrix(search32.Grid1.Row, 0)
    Unload search32
End If
Exit Sub
myerror:
End Sub

Private Sub cmdClient_Click()
publicFlag = 2
Clients.Show 1
End Sub
Private Sub cmdBarCode_Click()
If Not MYVALID Then Exit Sub
Dim tBarCode As New ADODB.Recordset
If Grid1.Rows = 1 Then Exit Sub
tBarCode.Open "addprint", CON, adOpenKeyset, adLockReadOnly, adCmdTable
tBarCode.Find "doc_no = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
If Not tBarCode.EOF Then
    If MsgBox("«·›« Ê—…  „  —ÕÌ·Â« „‰ ﬁ»· .. «·€«¡ «·«’‰«› «·„—Õ·… ··ÿ»«⁄…", vbOKCancel + vbDefaultButton2, " —ÕÌ· ··ÿ»«⁄…") = vbCancel Then
        Exit Sub
    End If
End If
With Grid1
CON.Execute "DELETE * FROM ADDPRINT WHERE DOC_NO = " & MyParn(xDoc_no.Text)
For I = 1 To Grid1.Rows - 2
    CON.Execute "Insert Into ADDPRINT(Doc_no,Item,Quant,isPrint) " & _
               " Values(" & _
               addstring(xDoc_no.Text) & "," & _
               addstring(.TextMatrix(I, 1)) & "," & _
               addvalue(.TextMatrix(I, 3)) & "," & _
               "TRUE" & _
               ")"
Next
End With
End Sub

Private Sub cmdBarCodeopen_Click()
Dream_Bar.Show 1
End Sub

Private Sub cmdCash_Click()
xVisa.Text = ""
xLate.Caption = ""
xCash.Text = TurnValue(Val(xTotalDis.Caption), 0, "")
xCash_LostFocus
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    'on Error GoTo MyError
    CON.BeginTrans
    ' Õ–› «·„” ‰œ
   CON.Execute "Delete * From " & cFile & " where Doc_No = " & MyParn(xDoc_no.Text)
   CON.Execute "Delete * From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_no.Text)
    
    ' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
'    CON.Execute "DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND FILE1_11.[TYPE] = " & MyParn(cItemmove)
    
    ' Õ–› Õ—ﬂ… «·⁄„Ì· «Ê «·„Ê—œ
'    CON.Execute "DELETE  * FROM " & cFileMove & " WHERE DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(cClientmove)
           
    CON.CommitTrans
    CardTable.Requery
    If CardTable.BOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_no.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    End If
End If
Exit Sub
myerror:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdExit_Click()
Unload Me
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
bAddnew = True
myDefine
On Error Resume Next
xCode.SetFocus

End Sub
Private Sub cmdSave_Click()
'If myPublic = 0 Then If Not nofoundOther Then Exit Sub
If Not MYVALID Then Exit Sub
CalcTotals
If Not MyReplace Then Exit Sub
CardTable.Requery
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If bAddnew Then
    CmdNewInv_Click
Else
    CardTable.Find "Doc_No = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
    Handlecontrols LoadMode
    MyLoad
End If
End Sub
Private Sub CmdUndo_Click()
bAddnew = False
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
MyLoad
End Sub
Private Sub Cmditem_Click()
Dim bEditLocal As Boolean
bEditLocal = bEdit: bEdit = True
items.Show 1
bEdit = bEditLocal
End Sub

Private Sub cmdVisa_Click()
xCash.Text = ""
xLate.Caption = ""
xVisa.Text = TurnValue(Val(xTotalDis.Caption), 0, "")
xVisa_LostFocus
End Sub

Private Sub Command2_Click()
doprint 1
End Sub

Private Sub Command3_Click()
'If LCase(InputBox("«œŒ· ﬂ·„… «·”—!", "ﬂ·„… «·”—")) = "mor2008" Then
'    bRetvalue = False
'    PassWord2.Caption = "ﬂ·„… ”— «ŸÂ«— «·„»Ì⁄«  «·ÌÊ„Ì…"
'    Set PassWord2.myForm = Me
'    PassWord2.Show 1
    TDaySal.bShowPrf = bopt1
    TDaySal.Show 1
'End If
End Sub

Private Sub Command1_Click()
doprint 0
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
cDefBox = retDef("file0_50")
'If DefGet(Me.Name, "CUST") = "TRUE" Then
'    cDefClient = GetDesca("Select Min(Code) From file3_10")
'Else
cDefClient = retDef("file3_10", , "cash")
'End If
cDefClientDesca = retDef("file3_10", "Desca", "cash")
bEdit = True
Select Case myPublic
Case 0
    cCodeDesca = "«·⁄„Ì·"
    cFile = "File6_20"
    cFileHeader = "File6_20H"
    cFileClient = "File3_10"
    cFileMove = "File3_11"
    cFieldItem = "[IN]"
    cFieldClient = "[SAL]"
    cMoveName = "„»Ì⁄« "
    Me.Caption = "›« Ê—… „»Ì⁄« "
Case 1
    cCodeDesca = "«·⁄„Ì·"
    cFile = "FILE6_10"
    cFileHeader = "File6_10H"
    cFileClient = "File3_10"
    cFileMove = "File3_11"
    cFieldItem = "[IN]"
    cFieldClient = "[SAL]"
    cMoveName = "„—œÊœ „»Ì⁄« "
    lblClient.Caption = "«·⁄„Ì· :"
    Me.Caption = "›« Ê—… „—œÊœ „»Ì⁄« "
End Select
ItemTable.Open "file1_10", CON, adOpenStatic, adLockReadOnly, adCmdTable

Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT " & cFileHeader & ".*,FILE3_10.DESCA AS CLIENTDESCA FROM " & cFileHeader & _
               " LEFT JOIN FILE3_10 ON " & cFileHeader & ".Code = FILE3_10.CODE " & _
               " ORDER BY DOC_NO", CON, adOpenKeyset, adLockReadOnly, adCmdText

GRDTABLE.Open "Select * From " & cFile & " order by Row", CON, adOpenForwardOnly, adLockReadOnly, adCmdText

Set ClientTable = New ADODB.Recordset
ClientTable.Open cFileClient, CON, adOpenKeyset, adLockReadOnly, adCmdTable

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = CON.ConnectionString
DATA2.RecordSource = "FILE6_25"
Set xman.RowSource = DATA2
xman.ListField = "Desca"
xman.BoundColumn = "Code"
xman.BoundText = retDef("FILE6_25")

data4.ConnectionString = CON.ConnectionString
data4.RecordSource = "SELECT * FROM FILE0_50"
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

With Grid1
    .Cols = 9
    .Rows = 2
    .Editable = flexEDKbdMouse
End With

Set Grid1.DataSource = data3
data3.ConnectionString = CON.ConnectionString

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
    CmdNewInv_Click
Else
    myDefine
    FixGrd
    xDoc_no.Text = RetZero("1", 6)
End If
'grid1.ColHidden(grid1.Cols - 1) = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetKbLayout Lang_AR
GRDTABLE.Close
ItemTable.Close
Set GRDTABLE = Nothing
Set ItemTable = Nothing
On Error Resume Next
Unload Search3
Unload SEARCH31
Unload search32
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Grid1.Col = 1 Then
    GrdDesc Row
End If
CalcTotals

If Grid1.TextMatrix(Row, 1) = "" Then
'    grid1.RemoveItem Row
'    CalcTotals True
'    bInsert = False
    MakeSerial
    Exit Sub
End If

If Not validRows(Row) Then Exit Sub
If Not validRows Then
    DelValid
    bInsert = True
End If

If Grid1.TextMatrix(Row, 2) <> "" Then
    If Not validHeader Then
        MsgBox "·‰ Ì „ Õ›Ÿ «·›« Ê—… "
        Grid1.Cell(flexcpForeColor, Row, 0, Row, Grid1.Cols - 1) = vbRed
        Exit Sub
    End If
    If MyReplaceItem(Row) Then
        Grid1.Cell(flexcpForeColor, Row, 0, Row, Grid1.Cols - 1) = vbBlack
        If xDoc_no.Tag = DefineMode Then
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
            Handlecontrols LoadMode
        End If
        xbalanceitem.Caption = ""
    Else
        Grid1.Cell(flexcpForeColor, Row, 0, Row, Grid1.Cols - 1) = vbRed
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
If Grid1.Col = 0 Or Grid1.Col = 2 Or Grid1.Col = 4 Or Grid1.Col = 5 Or Grid1.Col = 7 Or Grid1.Col = 8 Then
    Grid1.Editable = flexEDNone
ElseIf Grid1.Col = 1 Then
    Grid1.Editable = IIf(validHeader, flexEDKbdMouse, flexEDNone)
    SetKbLayout IIf(Grid1.Col = 1, Lang_EN, Lang_AR)
    'Grid1.EditCell
Else
   Grid1.Editable = IIf(Trim(Grid1.TextMatrix(Grid1.Row, 1)) <> "" And validHeader, flexEDKbdMouse, flexEDNone)
   SetKbLayout IIf(Grid1.Col = 1, Lang_EN, Lang_AR)
   'Grid1.EditCell
End If
End Sub
Private Sub Grid1_GotFocus()
If Grid1.Row = 0 Then Grid1.Select 1, 1
Grid1_EnterCell
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Or (KeyCode = 13 And Shift = 2) Then xDiscount.SetFocus
If KeyCode = 45 And Grid1.Row <> Grid1.Rows - 1 Then
    Grid1.AddItem "", Grid1.Row
    MakeSerial
    bInsert = True
End If
If KeyCode = 112 Then
    If Grid1.Col = 1 And Grid1.Row <> 0 Then ItemsLookup
End If
End Sub

Private Sub grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Grid1.Col
        Case 1
            If Grid1.Row = Grid1.Rows - 1 Then
                Grid1.TextMatrix(Grid1.Rows - 1, 3) = ""
                Grid1.AddItem ""
                Grid1.Select Grid1.Rows - 2, 3
                MakeSerial
            Else
                Grid1.TextMatrix(Grid1.Row, 3) = ""
                Grid1.Select Grid1.Row, 3
            End If
        Case 3
            Grid1.Select Grid1.Rows - 2, 6
        Case 6
            Grid1.Select Grid1.Rows - 1, 1
    End Select
    CalcTotals
End If
End Sub
Private Sub grid1_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Grid1.Row = Grid1.Rows - 1 Then
    Grid1.AddItem ""
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
    MakeSerial
End If
If Col = 3 And Grid1.TextMatrix(Row, 1) <> "" Then
    cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(Grid1.TextMatrix(Row, 1)) & _
             " and Store = " & MyParn(xStore.BoundText) & " and Date <= " & DateSq(xdate.Text)
    xbalanceitem.Caption = GetDesca(cString)
End If

End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If GetDesca("select item from file1_10 where item = " & MyParn(Grid1.EditText)) = "" Then
        If Grid1.EditText <> "" Then MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
    If myPublic = 0 Then
        cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(Grid1.EditText) & _
             " and Store = " & MyParn(xStore.BoundText) & " and Date <= " & DateSq(xdate.Text)
        xbalanceitem.Caption = GetDesca(cString)
        
        If Val(xbalanceitem.Caption) <= 0 Then
            Cancel = True
        End If
    End If
End If
If Grid1.Col = 3 And myPublic = 0 Then
   If Val(xbalanceitem.Caption) - Val(Grid1.EditText) < 0 Then
        Cancel = True
   End If
End If
End Sub
Private Sub XBOX_Click(Area As Integer)
If Not xDoc_no.Enabled Then updateHeader
End Sub
Private Sub xBox_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xBox_LostFocus()
xCash.Enabled = (Trim(xBox.BoundText) <> "")
cmdCash.Enabled = (Trim(xBox.BoundText) <> "")
CalcTotals
If Not xBox.MatchedWithList Then
    xBox.BoundText = ""
    xCash.Text = ""
    xLate.Caption = Val(xTotalDis.Caption) - Val(xVisa.Text)
End If
If Not xDoc_no.Enabled Then updateHeader
xBox.BackColor = &H80000005
End Sub
Private Sub xCash_LostFocus()
CalcLate xCash
'CalcTotals
If Not xDoc_no.Enabled Then updateHeader
End Sub

Private Sub xCode_DblClick()
CLIENTLOOKUP
End Sub

Private Sub xCODE_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
xcodedesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
xcodedesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xCode.Text)) & ""
If Not xDoc_no.Enabled Then updateHeader
'xBalance.Caption = Format(GetDesca("Select sum(Format(val(SAL & '') - val (pay & ''),'Fixed')) FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
End Sub

Private Sub xCode_Validate(Cancel As Boolean)
If Trim(xCode.Text) = "" Then Cancel = True
End Sub

Private Sub xDate_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xDate_LostFocus()
If Not xDoc_no.Enabled Then updateHeader
xdate.BackColor = &H80000005
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xdate.Text) Then Cancel = True
End Sub

Private Sub xDiscount_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xDiscount_LostFocus()
xDiscount.BackColor = &H80000005
CalcTotals
updateHeader
End Sub
Private Function MYVALID() As Boolean
Dim I As Integer
If xDoc_no.Text = "" Then
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

If xcodedesca.Caption = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With Grid1
'If Not validRows(, False, True) Then
    'MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
'    Exit Function
'End If
DelValid , True
End With
MYVALID = True
End Function
Private Sub MyLoad(Optional bLeaveBal As Boolean = False)
xDoc_no.Text = CardTable!doc_no
xdate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!Store & ""
xBox.BoundText = CardTable!Box & ""
xman.BoundText = CardTable!MAN & ""
xNotes.Text = CardTable!NOTES & ""
xCode.Text = CardTable!CODE & ""
xcodedesca.Caption = CardTable!ClientDesca & ""
xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = TurnValue(Val(CardTable!Discount & ""), 0, "")
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")
xCash.Text = TurnValue(Val(CardTable!Cash & ""), 0, "")
xVisa.Text = TurnValue(Val(CardTable!Visa & ""), 0, "")
If Not bLeaveBal Then xBALANCE.Caption = ""
With Grid1
    cField1 = "iif(val(" & cFile & ".Discount & '') = 0,Null,val(" & cFile & ".Discount & ''))"
    cString = "SELECT " & cFile & ".ROW, " & cFile & ".ITEM, FILE1_10.DESCA, Quant,Format(" & cFile & ".Price,'fixed'),0 AS Expr1," & cField1 & ",FILE1_10.PACKAGE,FILE1_10.UNIT " & _
          " FROM " & cFile & " LEFT JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_no.Text) & " ORDER by " & cFile & ".ROW"
    data3.RecordSource = cString
    data3.Refresh
    Grid1.AddItem ""
    MakeSerial
End With
Handlecontrols LoadMode
CalcTotals True
FixGrd
End Sub
Private Sub myDefine()
If CardTable.EOF And CardTable.BOF Then
    xDoc_no.Text = RetZero("1")
Else
    CardTable.MoveLast
    xDoc_no.Text = RetZero(Val(CardTable!doc_no) + 1, 6)
End If
xdate.Text = Format(Date, "dd-mm-yyyy")
xBALANCE.Caption = ""
xBox.BoundText = cDefBox
xCode.Text = cDefClient
xcodedesca.Caption = cDefClientDesca
xDiscount.Text = ""
xTotalDisItem.Caption = ""
xDisItem.Caption = ""
xtotal.Caption = ""
xTax.Text = ""
xLate.Caption = ""
xVisa.Text = ""
xCash.Text = ""
xbalanceitem.Caption = ""
'If xman.BoundText = "" Then xman.BoundText = cDefCasher
'xman.BoundText = cSalesMan
xTotalItem.Caption = ""
xTotalDis.Caption = ""
xusername.Text = ""
xNotes.Text = ""
xtotalQuant.Caption = ""
xRate.Text = ""
Grid1.Rows = 1
Grid1.AddItem ""
Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
'xDiscount.Enabled = (nMode = LoadMode)
'xTax.Enabled = (nMode = LoadMode)
xDoc_no.Enabled = (nMode = DefineMode)
xCash.Enabled = (Trim(xBox.BoundText) <> "")
cmdCash.Enabled = (Trim(xBox.BoundText) <> "")
xDoc_no.Tag = nMode
'xRate.Enabled = (nMode = LoadMode)
'xCash.Enabled = (nMode = LoadMode)
'xVisa.Enabled = (nMode = LoadMode)
End Sub

Private Sub xDoc_No_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_no.BackColor = &H80000005
xDoc_no.Text = RetZero(xDoc_no.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad True
End Sub
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Grid1.Row <> Grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        bValid = validRows(Grid1.Row, True)
        If bValid Then RemoveItem (Grid1.Row)
        Grid1.RemoveItem Grid1.Row
        CalcTotals
        updateHeader
        MakeSerial Grid1.Row
    End If
End If
'If KeyCode = 27 Then xDate.SetFocus
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'Select Case Grid1.Col
'    Case 1
'        If KeyCode = 27 Then
'            Exit Sub
'        End If
'        If KeyCode = 112 Then
'            ItemsLookup
'        End If
'End Select
End Sub
Private Sub GrdDesc(Row)
Grid1.TextMatrix(Row, 7) = ""
If Grid1.TextMatrix(Row, 1) = "" Then Exit Sub

ItemTable.Find "item = " & MyParn(Grid1.TextMatrix(Row, 1)), , adSearchForward, adBookmarkFirst
'ItemTable.Seek Array(Grid1.TextMatrix(Row, 1)), adSeekFirstEQ
If Not ItemTable.EOF Then
    Grid1.TextMatrix(Row, 2) = ItemTable!Desca
    If GetBoolean("select cust from file3_10 where code = " & MyParn(xCode.Text)) = 0 Then
        Grid1.TextMatrix(Row, 4) = ItemTable!price & ""
        Grid1.TextMatrix(Row, 6) = ItemTable!Discount & ""
    Else
        Grid1.TextMatrix(Row, 4) = ItemTable!price2 & ""
    End If
    Grid1.TextMatrix(Row, 7) = ItemTable!package & ""
    Grid1.TextMatrix(Row, 8) = ItemTable!unit & ""
End If
CalcTotals
End Sub
Private Function CalcTotals(Optional bCalcLate As Boolean)
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalDis As Double
With Grid1
For I = 1 To Grid1.Rows - 1
    nTotalitem = nTotalitem + (Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)))
    nDiscount = 1 - (Val(.TextMatrix(I, 6)) / 100)
    Grid1.TextMatrix(I, 5) = TurnValue(Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount, 0, "")
    nTotalDisItem = nTotalDisItem + (Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount)
    nTotalQuant = nTotalQuant + Val(Grid1.TextMatrix(I, 3))
Next
nDisItem = nTotalitem - nTotalDisItem
nTotalDis = nTotalDisItem - Val(xDiscount.Text)
nTotal = nTotalDis + Val(xTax.Text)
If nTotalitem <> 0 Then
    xRateDis.Text = TurnValue(Round((nTotalDisItem - nTotalDis) / nTotalDisItem * 100, 2), 0, "")
End If
xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
xTotalItem.Caption = Format(nTotalitem, "Fixed")
xTotalDisItem.Caption = Format(nTotalDisItem, "Fixed")
xDisItem.Caption = Format(nDisItem, "Fixed")
xTotalDis.Caption = Format(nTotalDis, "Fixed")
xtotal.Caption = Format(nTotal, "Fixed")
xtotalQuant.Caption = Format(nTotalQuant, "#0.0000")
'If Trim(xBox.BoundText) = "" Then
'    xCash.Text = ""
'    xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xVisa.Text), 0, "")
'Else
If bCalcLate Or xBox.BoundText = "" Then
    xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xCash.Text) - Val(xVisa.Text), 0, "")
Else
    If Val(xLate.Caption) = 0 Then
        xCash.Text = TurnValue(Val(xTotalDis.Caption) - Val(xVisa.Text), 0, "")
        xLate.Caption = ""
    Else
        xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xCash.Text) - Val(xVisa.Text), 0, "")
    End If
End If
'End If

'If Val(xLate.Text) < 0 Then
    
'End If
'If Val(xVisa.Text) < 0 Then
'    xVisa.Text = ""
'    If Val(xLate.Caption) <> 0 Then
'        xLate.Caption = Val(xLate.Caption) - Abs(Val(xVisa.Text))
'    Else
'
'    End If
'End If
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Format([DATE],'yyyy/mm/dd'), " & cFileClient & ".Desca " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE )"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ " & cCodeDesca
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load SEARCH31
SEARCH31.Caption = "«” ⁄·«„"
SEARCH31.Show 1
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
    nRow = FoundOtherRow(I, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & Grid1.TextMatrix(nRow, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_no.Text = "" Then Cancel = True


End Sub

Private Sub xMAN_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xman_LostFocus()
If Not xDoc_no.Enabled Then updateHeader
xman.BackColor = &H80000005
End Sub

Private Sub xMAN_Validate(Cancel As Boolean)
If Not xman.MatchedWithList Then xman.BoundText = ""
If Trim(xman.BoundText) = "" Then Cancel = True
End Sub

Private Sub xNotes_LostFocus()
If Not xDoc_no.Enabled Then updateHeader
End Sub

Private Sub xRate_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xRate_LostFocus()
xRate.BackColor = &H80000005
If Val(xRate.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
updateHeader
End Sub
Private Function RetItemBalance(citem, cStore, dDate) As Double
If citem = "" Then Exit Function
movetable.Seek Array(citem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> citem Or cStore <> movetable!Store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    'If Not (movetable!Type = cItemmove And movetable!Doc_Id = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    'End If
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub FixGrd()
With Grid1
.FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ì|" & "«·Œ’„|" & "«·⁄»Ê…|" & "«·ÊÕœ…"
.ColWidth(0) = 500
.ColWidth(1) = 1800
.ColWidth(2) = 4500
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From " & cFileClient
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub

Private Sub xRateDis_Lostfocus()
xDiscount.Text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.Text) / 100))
CalcTotals
updateHeader
End Sub

Private Sub xStore_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xstore_LostFocus()
xStore.BackColor = &H80000005
If Not xDoc_no.Enabled Then updateHeader
End Sub

Private Sub xStore_Validate(Cancel As Boolean)
If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub

Private Sub xTax_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xTax_LostFocus()
xTax.BackColor = &H80000005
CalcTotals
updateHeader
End Sub
Private Function MyReplaceItem(nRow) As Boolean
If bInsert Then
    If MyReplace Then
        MyReplaceItem = True
        bInsert = False
        Exit Function
    End If
End If

Dim nTry As Integer
On Error Resume Next
CON.BeginTrans
For nTry = 1 To 10
    If xDoc_no.Tag = DefineMode Then
        CON.Execute "insert into " & cFileHeader & "(doc_no,[date],code,store,box,man,totalItem,Discount,Tax,cash,visa,total,Notes,username) values(" & _
                    addstring(xDoc_no.Text) & "," & _
                    DateSq(xdate.Text) & "," & _
                    addstring(xCode.Text) & "," & _
                    addstring(xStore.BoundText) & "," & _
                    addstring(xBox.BoundText) & "," & _
                    addstring(xman.BoundText) & "," & _
                    Val(xTotalItem.Caption) & "," & _
                    Val(xDiscount.Text) & "," & _
                    Val(xTax.Text) & "," & _
                    Val(xCash.Text) & "," & _
                    Val(xVisa.Text) & "," & _
                    addvalue(xtotal.Caption) & "," & _
                    addstring(xNotes.Text) & "," & _
                    addstring(sUserName) & _
                    ")"
    Else
        updateHeader
    End If
    If Err.Number = 0 Then
        With Grid1
                CON.Execute "Insert Into " & cFile & " (Doc_no,Item,Quant,Price,Total,Discount,row,Cost) " & _
                            " Values(" & _
                            addstring(xDoc_no.Text) & "," & _
                            addstring(.TextMatrix(nRow, 1)) & "," & _
                            addvalue(.TextMatrix(nRow, 3)) & "," & _
                            addvalue(.TextMatrix(nRow, 4)) & "," & _
                            Val(.TextMatrix(nRow, 3)) * Val(.TextMatrix(nRow, 4)) & "," & _
                            Val(.TextMatrix(nRow, 6)) & "," & _
                            nRow & "," & _
                            Val(itemCost(Grid1.TextMatrix(nRow, 1), xdate.Text)) & _
                            ")"
                If Err.Number = -2147467259 Then
                    Err.Clear
                    CON.Execute "update " & cFile & _
                                " set item = " & addstring(.TextMatrix(nRow, 1)) & _
                                ", Quant = " & Val(.TextMatrix(nRow, 3)) & _
                                ", Price = " & Val(.TextMatrix(nRow, 4)) & _
                                ", Total = " & Val(.TextMatrix(nRow, 3)) * Val(.TextMatrix(nRow, 4)) & _
                                ", discount = " & Val(.TextMatrix(nRow, 6)) & _
                                ",Cost = " & Val(itemCost(Grid1.TextMatrix(nRow, 1), xdate.Text)) & _
                                " where Doc_no = " & MyParn(xDoc_no.Text) & " and Row = " & nRow, nRecord
                End If
            If Err.Number <> 0 Then GoTo myerror
        End With
    End If
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 Then
        Err.Clear
        xDoc_no.Text = RetZero(Val(xDoc_no.Text) + 1)
    End If
    If Err.Number <> 0 Then GoTo myerror
Next
CON.CommitTrans
bInsert = False
MyReplaceItem = True
Exit Function
myerror:
prog1.Visible = False
CON.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Private Function RemoveItem(nRow) As Boolean
On Error GoTo myerror
CON.BeginTrans
CON.Execute "Delete * From " & cFile & " where Doc_No = " & MyParn(xDoc_no.Text) & " and row = " & nRow
For I = nRow + 1 To Grid1.Rows - 2
    CON.Execute "update " & cFile & " set row = " & (I - 1) & " where row = " & I & " and doc_no = " & MyParn(xDoc_no.Text)
Next
CON.CommitTrans
updateHeader
Exit Function
myerror:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function updateHeader() As Boolean
CON.BeginTrans
CON.Execute "update " & cFileHeader & " set " & _
            "[DATE] = " & DateSq(xdate.Text) & _
            ",BOX = " & addstring(xBox.BoundText) & _
            ",CODE = " & addstring(xCode.Text) & _
            ",store = " & addstring(xStore.BoundText) & _
            ",man = " & addstring(xman.BoundText) & _
            ",Notes = " & addstring(xNotes.Text) & _
            ",DISCOUNT = " & Val(xDiscount.Text) & _
            ",TAX = " & Val(xTax.Text) & _
            ",CASH = " & Val(xCash.Text) & _
            ",VISA = " & Val(xVisa.Text) & _
            ",TOTAL = " & addvalue(xtotal.Caption) & _
            ",username = " & addstring(sUserName) & _
            " where doc_no = " & MyParn(xDoc_no.Text)
CON.CommitTrans
Exit Function
myerror:
CON.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function validHeader(Optional bMsg As Boolean = True) As Boolean
If Trim(xDoc_no.Text) = "" Then
    If bMsg Then MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xdate.Text) Then
    If bMsg Then MsgBox "«· «—ÌŒ €Ì— ’«·Õ «Ê „”Ã·"
    Exit Function
End If

If Trim(xStore.BoundText) = "" Then
    If bMsg Then MsgBox "«·„Œ“‰ €Ì— „”Ã·"
    Exit Function
End If

If Trim(xCode.Text) = "" Then
    If bMsg Then MsgBox "ﬂÊœ «·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

validHeader = True
End Function

Private Sub xVisa_LostFocus()
'If Val(xLate.Caption) = 0 Then
'    xCash.Text = TurnValue(Val(xTotalDis.Caption) - Val(xVisa.Text), 0, "")
'Else
'    xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xCash.Text) - Val(xVisa.Text), 0, "")
'End If
CalcLate xVisa
'CalcTotals
If Not xDoc_no.Enabled Then updateHeader
End Sub
Private Sub doprint(NFLAG As Byte)
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To Grid1.Rows - 2
    temptable.AddNew
    If myPublic = 0 Then
        If NFLAG = 0 Then
            temptable!str21 = "≈–‰  ”·Ì„ »÷«⁄… "
        Else
            temptable!str21 = "≈–‰ ’—› »÷«⁄…"
        End If
    Else
        If NFLAG = 0 Then
            temptable!str21 = "≈–‰ „— Ã⁄ »÷«⁄… „‰ ⁄„Ì· "
        Else
            temptable!str21 = "≈–‰ „— Ã⁄ »÷«⁄… ··„Œ“‰"
        End If
    End If
    temptable!str1 = xDoc_no.Text
    temptable!str2 = xdate.Text
    temptable!str3 = Format(xCode.Text)
    temptable!str4 = xcodedesca.Caption
    temptable!str5 = xStore.Text
    temptable!STR6 = IIf(Val(xLate.Caption) = 0, "‰ﬁœÌ", "¬Ã·")
    temptable!Str11 = TurnValue(Grid1.TextMatrix(I, 1))
    
    temptable!str12 = TurnValue(Grid1.TextMatrix(I, 2))
    temptable!val1 = Val(Grid1.TextMatrix(I, 3))
    temptable!val2 = Val(Grid1.TextMatrix(I, 4))
    temptable!val3 = Val(Grid1.TextMatrix(I, 6))
    temptable!val4 = Val(Grid1.TextMatrix(I, 5))
    temptable!Val10 = I
    
    temptable!val5 = Val(xTotalItem.Caption)
    temptable!Val6 = Val(xDisItem.Caption)
    temptable!Val7 = Val(xDiscount.Text)
    temptable!Val8 = Val(xtotal.Caption)
    temptable!val9 = Val(xtotalQuant.Caption)
    
    temptable!str10 = MyOnly(Val(xtotal.Caption))
    temptable!val9 = myPublic
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If NFLAG = 0 Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\sales.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Order.rpt"
End If
Main.REPORT1.DataFiles(0) = tempPath
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub CalcLate(pControl)
xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xCash.Text) - Val(xVisa.Text), 0, "")
If Val(xLate.Caption) < 0 Then
    If pControl.Name = xCash.Name Then
        If Val(xVisa.Text) >= Abs(Val(xLate.Caption)) Then
            xVisa.Text = Val(xVisa.Text) - Abs(Val(xLate.Caption))
            xLate.Caption = 0
        Else
            xLate.Caption = Val(xVisa.Text) - Abs(Val(xLate.Caption))
            xVisa.Text = ""
        End If
    Else
        If Val(xCash.Text) >= Abs(Val(xLate.Caption)) Then
            xCash.Text = Val(xCash.Text) - Abs(Val(xLate.Caption))
            xLate.Caption = 0
        Else
            xLate.Caption = Val(xCash.Text) - Abs(Val(xLate.Caption))
            xCash.Text = ""
        End If
    End If
End If
End Sub
Private Function validRows(Optional prow = -1, Optional igMsg As Boolean = True, Optional bReqQuant As Boolean = False) As Boolean
For nRow = IIf(prow = -1, 1, prow) To IIf(prow = -1, Grid1.Rows - 2, prow)
    If Trim(Grid1.TextMatrix(nRow, 1)) = "" Then
        If Not igMsg Then MsgBox "«·’‰› ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã· "
        Exit Function
    End If
    If Val(Grid1.TextMatrix(nRow, 3)) = 0 And bReqQuant Then
        If Not igMsg Then MsgBox "«·ﬂ„Ì… ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã·… "
        Exit Function
    End If
Next
validRows = True
End Function
Private Function DelValid(Optional prow = -1, Optional bReqQuant As Boolean = False) As Boolean
For nRow = Grid1.Rows - 2 To 1 Step -1
    If (nRow = prow) Or prow = -1 Then
'        If grid1.TextMatrix(nRow, 1) = "" Then
'            grid1.RemoveItem nRow
'        End If
        If Not validRows(nRow, , True) Then Grid1.RemoveItem nRow
    End If
Next
MakeSerial
End Function

Sub myproc2(nDoc_no)
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    MyLoad
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
