VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form feesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√ﬂÊ«œ «·«’‰«›"
   ClientHeight    =   9285
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   14160
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14160
   Begin VB.Frame Frame2 
      Height          =   9195
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   6945
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8970
         Left            =   45
         TabIndex        =   22
         Top             =   135
         Width           =   6855
         _cx             =   12091
         _cy             =   15822
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
         Cols            =   6
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
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   675
      Width           =   6990
      Begin VB.TextBox xQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1350
         Width           =   2940
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   990
         Width           =   300
      End
      Begin VB.TextBox xClient 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4050
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   1410
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xCar_type 
         Height          =   330
         Left            =   2520
         TabIndex        =   20
         Top             =   990
         Width           =   2940
         _ExtentX        =   5186
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
      Begin VB.Label xClient_Desca 
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   630
         Width           =   3840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
         DragMode        =   1  'Automatic
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1395
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·”Ì«—…"
         DragMode        =   1  'Automatic
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
         DragMode        =   1  'Automatic
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ"
         DragMode        =   1  'Automatic
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
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   390
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2970
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   11
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
         Picture         =   "fees.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "fees.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   855
         TabIndex        =   12
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
         Picture         =   "fees.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "fees.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   13
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
         Picture         =   "fees.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "fees.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   14
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
         Picture         =   "fees.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "fees.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6990
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
         Left            =   3510
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fees.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2325
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fees.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fees.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fees.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4590
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fees.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5760
         Picture         =   "fees.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   2070
      Top             =   180
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4770
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
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   330
      Left            =   2970
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
   Begin MSAdodcLib.Adodc data2 
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
   Begin VB.Label xRecord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   6990
   End
End
Attribute VB_Name = "feesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Public SCODE As String
Dim nRound As Long
Dim con As New ADODB.Connection
Dim cFilter As String, cFilterLookup As String, cList As String
Dim oSearch As New Search3, oSearchClient As New Search3
Dim formMode As Byte, osearchitem As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdCont_Click()
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sTable = "Container_codes"
oFlagfrm.sCaption = "‰Ê⁄ «·Õ«ÊÌ…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
myloadgrd2
End Sub
Private Sub cmdCarType_Click()
Dim oFlagfrm As New flag_mainfrm, SCODE As String
SCODE = xGroup.BoundText
oFlagfrm.sTable = "CAR_TYPE_CODES"
oFlagfrm.sCaption = "‰Ê⁄ «·”Ì«—…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data2.Refresh
If SCODE <> "" Then xCar_type.BoundText = SCODE
If Not xCar_type.MatchedWithList Then xCar_type.BoundText = ""
End Sub
Private Sub Form_Activate()
If SCODE <> "" Then
    On Error Resume Next
    xCode.SetFocus
    Err.Clear
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Name <> xCost_sup.Name Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
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
nRound = 2
bEdit = True
openCon con

DATA1.ConnectionString = strCon
DATA1.RecordSource = "PLACE_CODES"
Set xFrom.RowSource = DATA1
xFrom.ListField = "Desca"
xFrom.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "CAR_TYPE_CODES"
Set xCar_type.RowSource = data2
xCar_type.ListField = "Desca"
xCar_type.BoundColumn = "Code"

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

cList = StrList("SELECT CODE,DESCA FROM PLACE_CODES ORDER BY DESCA")
openCardTable
myUndo
If SCODE = "" Then CmdAdd_Click
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.Text = ""
On Error Resume Next
xClient.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From Fair  Where CODE = " & xCode.Text
    con.Execute "Delete  From Fair_sub  Where CODE = " & xCode.Text
    con.CommitTrans
    If SCODE <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "ITEM < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
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
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If SCODE <> "" Then
    Unload Me
    Exit Sub
End If
CheckCont
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
xClient.SetFocus
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
cardLookupAll Me, oSearch, cFilter, True
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
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode) And Trim(SCODE) = ""
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xClient.Text = ""
xClient_Desca.Caption = ""
xCar_type.BoundText = ""
xQuant.Text = ""
grid1.Rows = 1
MyAddItem
Handlecontrols DefineMode
xRecord.Caption = "«÷«›… ”Ã· ÃœÌœ"
End Sub
Sub myload()
xCode.Text = CardTable!code & ""
xClient_Desca.Caption = CardTable!client_Desca & ""
xCar_type.BoundText = CardTable!car_type & ""
xQuant.Text = Myvalue(CardTable!Quant)
myloadgrd
xRecord.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
End Sub
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CLIENT", addstring(xClient.Text))
aInsert = AddFlag(aInsert, "[CAR_TYPE]", addstring(xCar_type.BoundText))
aInsert = AddFlag(aInsert, "[QUANT]", Val(xQuant.Text))
On Error GoTo myerror
con.BeginTrans
If xCode.Text = "" Then
    xCode.Text = RetZero(Val(Newflag("FAIR", "ITEM")))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "FAIR")
Else
    con.Execute addUpdate(aInsert, "FAIR", "ITEM = " & addstring(xCode.Text))
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
If ActiveControl.Name = CmdInform.Name Then
    xCode.Text = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
    osearchitem.Hide
    myUndo
ElseIf ActiveControl.Name = xClient.Name Then
    xClient.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    SendKeys "{TAB}"
    oSearchClient.Hide
    Unload oSearchClient
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload oSearch
Set oSearch = Nothing
Err.Clear
closeCon con
UnloadAllForms "search3"
Unload itemsfrm
Set itemsfrm = Nothing
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And grid1.Row = grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, 1) = "" And grid1.Col = 0 Then
    KeyAscii = 0
    If cmdSave.Enabled Then
        cmdSave_Click
        CmdAdd_Click
    End If
End If
End Sub
Private Sub grid2_GotFocus()
If grid2.Row <= 0 And grid2.Rows > 1 Then grid2.Select 1, 2
End Sub
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cellpos2 KeyCode, grid2.Row, grid2.Col
End Sub

Private Sub grid2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then cellpos2 KeyCode, Row, Col
End Sub

Private Sub xClient_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchClient
End Sub
Private Sub xClient_LostFocus()
myLostFocus xClient
xClient_Desca.Caption = ""
If xClient.Text = "" Then Exit Sub
xClient.Text = RetZero(xClient.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file3_10 where code = " & MyParn(xClient.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xClient_Desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Sub xIgDiscount_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Calctotals
End Sub
Private Sub XCODE_LostFocus()
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "CODE = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    If xCode.Tag = LoadMode Then mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Trim(xClient.Text) = "" Then
    If Not bIgMsg Then MsgBox "«·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

If Not bIgMsg Then
    If IsEmpty(GetField("select code from file3_10 where code = " & MyParn(xClient.Text))) Then
        MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    End If
End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If ValidInt(xCode.Text) Then
        CardTable.Find "CODE = " & xCode.Text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT FAIR.*,FILE3_10.DESCA AS CODE_DESCA FROM FAIR INNER JOIN FILE3_10 ON FAIR.CLIENT = FILE3_10.CODE"
cFilter = ""
If SCODE <> "" Then cFilter = cFilter & turn(cFilter, " and ") & " CODE = " & MyParn(SCODE)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY FAIR.[ITEM]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myRemove(Row As Long)
grid.RemoveItem Row
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEdit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM FEE_SUB WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 And aEditRecord And (grid1.Col = 0 Or grid1.Col = 1) Then
    Places_LookupAll Me, osearchPlace
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
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then MyAddItem
Calctotals
If myreplace(Row) Then
    Handlecontrols LoadMode
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
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Val(.TextMatrix(Row, 2)) = 0 Then Exit Function
If Val(.TextMatrix(Row, 3)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        grid1.RemoveItem OldRow
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
If aEditRecord Then
    grid1.Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Or Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        If IsNumeric(.Cell(flexcpTextDisplay, Row, Col)) Then
            aRet = GetFields("Select code,desca from file0_50 where code = " & MyParn(RetZero(.EditText)))
            If Not IsEmpty(aRet) Then
                .EditText = retFlag(aRet, "desca")
            End If
        Else
            Cancel = True
            MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
            Exit Sub
        End If
    End If
ElseIf Col = 2 Then
    If (Not IsNumeric(grid1.EditText)) Then
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„‰|" & "≈·Ì|" & "«·‰Ê·Ê‰|"
.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 800
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColComboList(0) = cList
.ColComboList(1) = cList
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 4 Then
    .Col = Col + 1 + IIf(Col = 0, 1, 0)
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub MyAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "ITEM_MAIN", addstring(xCode.Text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "QUANT", Val(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "COST", Val(grid1.TextMatrix(i, 3)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_20")
        Else
            con.Execute addUpdate(aInsert, "FILE1_20", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
cString = "SELECT FILE1_20.ITEM,FAIR.DESCA,QUANT,FILE1_20.COST,FILE1_20.TOTAL,FILE1_20.ID " & _
          " FROM FILE1_20 INNER JOIN FAIR ON FILE1_20.ITEM = FAIR.ITEM"
cString = cString & turn(cString) & "ITEM_MAIN = " & MyParn(xCode.Text)
DATA11.RecordSource = cString
DATA11.Refresh
MyAddItem
Fixgrd
End With
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
Private Sub xPrice_LostFocus()
myLostFocus xPrice
If Val(xCost_price.Caption) <> 0 And Val(xPrice.Text) <> 0 Then
    If Round(Val(xRate.Text), nRound) <> Round((Val(xPrice.Text) / Val(xCost_price.Caption)) - 100, 2) Then
        xRate.Text = Round(((Val(xPrice.Text) / Val(xCost_price.Caption)) * 100) - 100, nRound)
    End If
Else
   xRate.Text = ""
End If
Calctotals
End Sub
Private Sub xPrice2_LostFocus()
myLostFocus xPrice2
If Val(xCost_price.Caption) <> 0 And Val(xPrice2.Text) <> 0 Then
    If Round(Val(xrate2.Text), nRound) <> Round((Val(xPrice2.Text) / Val(xCost_price.Caption)) - 100, 2) Then
        xrate2.Text = Round(((Val(xPrice2.Text) / Val(xCost_price.Caption)) * 100) - 100, nRound)
    End If
Else
   xrate2.Text = ""
End If
Calctotals
End Sub
Private Sub xrate2_GotFocus()
myGotFocus xrate2
End Sub
Private Sub xrate2_LostFocus()
myLostFocus xrate2
Calctotals
If Val(xCost_price.Caption) <> 0 And Val(xPrice2.Text) <> 0 Then
    If Round(Val(xrate2.Text), nRound) <> Round((Val(xPrice2.Text) / Val(xCost_price.Caption)) - 100, 2) Then
        xrate2.Text = Round(((Val(xPrice2.Text) / Val(xCost_price.Caption)) * 100) - 100, nRound)
    End If
Else
   xrate2.Text = ""
End If
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xRate_LostFocus()
Calctotals
If Val(xCost_price.Caption) <> 0 And Val(xPrice.Text) <> 0 Then
    If Round(Val(xRate.Text), nRound) <> Round((Val(xPrice.Text) / Val(xCost_price.Caption)) - 100, 2) Then
        xRate.Text = Round(((Val(xPrice.Text) / Val(xCost_price.Caption)) * 100) - 100, nRound)
    End If
Else
   xRate.Text = ""
End If
myLostFocus xRate
End Sub
Private Sub xMotive_GotFocus()
myGotFocus xMotive
End Sub
Private Sub xMotive_LostFocus()
myLostFocus xMotive
Calctotals
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
Calctotals
End Sub
Private Sub xCost_sup_GotFocus()
myGotFocus xCost_sup
End Sub
Private Sub xCost_sup_LostFocus()
myLostFocus xCost_sup
If xCode.Tag = DefineMode Then
    If Val(retFlag(aAddress, "FAIR")) <> 0 And Val(xCost_sup.Text) <> 0 Then
        xFair.Text = Round(Val(xCost_sup.Text) * (Val(retFlag(aAddress, "FAIR")) / 100), 2)
    End If
End If
Calctotals
End Sub
Private Sub xClient_GotFocus()
myGotFocus xClient
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xGroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xNOTES_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNOTES_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xFair_GotFocus()
myGotFocus xFair
End Sub
Private Sub xFair_LostFocus()
myLostFocus xFair
Calctotals
End Sub
Private Sub xPrice_GotFocus()
myGotFocus xPrice
End Sub
Private Sub xPrice2_GotFocus()
myGotFocus xPrice2

End Sub
Private Sub CheckCont()
If Not ItemField(xCode.Text, "item_fix", con) & "" Then Exit Sub
If grid1.Row > 2 Then Exit Sub
Dim loctable As New ADODB.Recordset
Set loctable = Item_fix_table(xCode.Text, con)
con.BeginTrans
Dim aInsert As Variant, nCost_sup As Double
Dim aitem As Variant
Do Until loctable.EOF
    con.Execute "UPDATE FILE1_20 SET FILE1_20.COST = FAIR.COST FROM FILE1_20 INNER JOIN FAIR ON FILE1_20.ITEM = FAIR.ITEM WHERE FILE1_20.ITEM_MAIN = " & MyParn(loctable!Item_MAIN)
    
    aitem = ItemFields(loctable!Item_MAIN, con)
    nCost_sup = Val(GetField("Select sum(FILE1_20.TOTAL) FROM FILE1_20 WHERE ITEM_MAIN = " & MyParn(loctable!Item_MAIN), con))
    If xIgDiscount.Value = 0 Then
        nCost_Net = nCost_sup * (1 - (Val(retFlag(aitem, "DISCOUNT") & "") / 100))
    Else
        nCost_Net = nCost_sup
    End If
    ncost_charge = Val(GetField("Select sum(ITEM_CHARGE.VALUE) FROM ITEM_CHARGE WHERE ITEM =   " & MyParn(loctable!Item_MAIN), con) & "")
    nCost = nCost_Net + ncost_charge + Val(retFlag(aitem, "Motive")) + Val(retFlag(aitem, "fair"))
    aInsert = AddFlag(Empty, "COST_SUP", nCost_sup)
    aInsert = AddFlag(aInsert, "COST_NET", nCost_Net)
    aInsert = AddFlag(aInsert, "COST", nCost)
    con.Execute addUpdate(aInsert, "FAIR", "ITEM = " & MyParn(loctable!Item_MAIN))
    loctable.MoveNext
Loop
con.CommitTrans
item_cost_fixfrm.SCODE = xCode.Text
item_cost_fixfrm.sDesca = xDesca.Text
item_cost_fixfrm.Show 1
End Sub

