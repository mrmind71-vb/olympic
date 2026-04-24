VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form travelfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·—Õ·« "
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19065
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
   ScaleHeight     =   9915
   ScaleWidth      =   19065
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox xTrust_close 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«€·«ﬁ «·⁄Âœ…"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   14175
      RightToLeft     =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   8055
      Width           =   1455
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "≈€·«ﬁ «·„” ‰œ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   14175
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   7740
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Õ”«» «·‰Ê·Ê‰"
      Height          =   735
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   7695
      Width           =   13560
      Begin VB.CommandButton cmdCalcTotal 
         Caption         =   "..."
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label38 
         Caption         =   "≈Ã„«·Ì «·Ê“‰"
         Height          =   240
         Left            =   12330
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label xWeight_Total 
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
         Left            =   10845
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label40 
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   240
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   315
         Width           =   645
      End
      Begin VB.Label xTotal_weight 
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
         Height          =   390
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xExtend 
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
         Height          =   390
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label36 
         Caption         =   "«·“Ì«œ…"
         Height          =   240
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   315
         Width           =   600
      End
      Begin VB.Label xDiscount 
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Œ’„"
         Height          =   240
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   315
         Width           =   555
      End
      Begin VB.Label xWeight_value 
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
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label32 
         Caption         =   "≈Ã„«·Ì ﬁÌ„… «·Ê“‰"
         Height          =   240
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.TextBox xNotes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   90
      Width           =   4740
   End
   Begin VB.Frame Frame5 
      Caption         =   "„Ê—œÌ‰ «·Œœ„…"
      Height          =   1050
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   0
      Width           =   6675
      Begin VB.TextBox xTotal_sup 
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
         Left            =   4545
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   630
         Width           =   1230
      End
      Begin VB.TextBox xCode_sup 
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
         Left            =   4545
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "«·„” Õﬁ"
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
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ"
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
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   270
         Width           =   510
      End
      Begin VB.Label xCode_sup_desca 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   270
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   12555
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   -45
      Width           =   6405
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5295
         Picture         =   "travel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4245
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1095
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2145
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1050
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
         Left            =   3195
         MaskColor       =   &H00FFFFFF&
         Picture         =   "travel.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "⁄Âœ… «·”›—Ì…"
      Height          =   3435
      Left            =   9270
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   3105
      Width           =   9690
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   3075
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   9465
         _cx             =   16695
         _cy             =   5424
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
      Left            =   15795
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   7695
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   36
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
         Picture         =   "travel.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "travel.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   37
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
         Picture         =   "travel.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "travel.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   38
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
         Picture         =   "travel.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "travel.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   39
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
         Picture         =   "travel.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "travel.frx":1CE87
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
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6570
      Width           =   6990
      Begin VB.Label xRest_trust 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   585
         Width           =   195
      End
      Begin VB.Label Label19 
         Caption         =   "≈Ã„«·Ì «·⁄Âœ…"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Width           =   1710
      End
      Begin VB.Label xProfit_rate 
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "«·—»Õ"
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
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   585
         Width           =   1305
      End
      Begin VB.Label xProfit 
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "«·»«ﬁÌ"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   585
         Width           =   630
      End
      Begin VB.Label xTotal_trust 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label xTotal_cost 
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "≈Ã„«·Ì «·„’—Ê›"
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
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1350
         Visible         =   0   'False
         Width           =   165
      End
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
      Height          =   2400
      Left            =   6795
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   675
      Width           =   12165
      Begin VB.CommandButton cmdPlace 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox xGas 
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Tag             =   "N"
         Top             =   1980
         Width           =   2805
      End
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3870
         MaxLength       =   300
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1980
         Width           =   7215
      End
      Begin VB.CommandButton cmdCargo 
         Caption         =   "..."
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xClass 
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Tag             =   "new"
         Top             =   1260
         Width           =   2805
      End
      Begin VB.TextBox xWeight 
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Tag             =   "new"
         Top             =   900
         Width           =   2805
      End
      Begin VB.TextBox xTotal 
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
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Tag             =   "N"
         Top             =   1620
         Width           =   2805
      End
      Begin VB.TextBox xCar 
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
         Left            =   10440
         MaxLength       =   5
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1620
         Width           =   645
      End
      Begin VB.TextBox xDistance 
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
         Left            =   3870
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1260
         Width           =   2040
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
         Left            =   9990
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
         Left            =   9990
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
         Left            =   3825
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   2085
      End
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   6885
         TabIndex        =   3
         Top             =   900
         Width           =   4200
         _ExtentX        =   7408
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
      Begin MSDataListLib.DataCombo xDriver2 
         Height          =   330
         Left            =   6885
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1260
         Width           =   4200
         _ExtentX        =   7408
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
      Begin MSDataListLib.DataCombo xCargo 
         Height          =   330
         Left            =   495
         TabIndex        =   11
         Tag             =   "new"
         Top             =   180
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
      Begin MSDataListLib.DataCombo xFollower 
         Height          =   330
         Left            =   135
         TabIndex        =   13
         Tag             =   "new"
         Top             =   540
         Width           =   2805
         _ExtentX        =   4948
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
      Begin MSDataListLib.DataCombo xPlace1 
         Height          =   330
         Left            =   4230
         TabIndex        =   7
         Top             =   540
         Width           =   1680
         _ExtentX        =   2963
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
      Begin MSDataListLib.DataCombo xPlace2 
         Height          =   330
         Left            =   4230
         TabIndex        =   8
         Top             =   900
         Width           =   1680
         _ExtentX        =   2963
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
      Begin MSDataListLib.DataCombo xTrailer 
         Height          =   330
         Left            =   3870
         TabIndex        =   10
         Top             =   1620
         Width           =   2040
         _ExtentX        =   3598
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
      Begin VB.Label Label34 
         Caption         =   "«·„ﬁÿÊ—…"
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
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "”Ê·«—"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
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
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label22 
         Caption         =   "«·›∆…"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label Label21 
         Caption         =   "«·ﬂ„Ì…"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label15 
         Caption         =   "«· »«⁄"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Label13 
         Caption         =   "«·Õ„Ê·…"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   225
         Width           =   840
      End
      Begin VB.Label xCar_type 
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
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1620
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "”«∆ﬁ À«‰"
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
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "«·‰Ê·Ê‰"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label xCar_Desca 
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1620
         Width           =   1680
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "«·”Ì«—…"
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
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1620
         Width           =   570
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„”«›…"
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1350
         Width           =   570
      End
      Begin VB.Label Label6 
         Caption         =   "Õ Ï"
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   855
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "„‰"
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
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   525
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
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Width           =   3075
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·”«∆ﬁ"
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
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   900
         Width           =   540
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   28
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
         Left            =   11205
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   885
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
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   540
         Width           =   480
      End
   End
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   360
      Top             =   8775
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
      Caption         =   "„’—Ê› «·—Õ·…"
      Height          =   3390
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   3105
      Width           =   9195
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   2985
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   8970
         _cx             =   15822
         _cy             =   5265
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
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   6570
      Width           =   8610
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   270
         Width           =   195
      End
      Begin VB.Label xGasPerKilo_Differ 
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label33 
         Caption         =   "„⁄œ· «·«‰Õ—«›"
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label xGasPerKilocar 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "„⁄œ· «” Â·«ﬂ „› —÷"
         Height          =   285
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label xGascar 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "«” Â·«ﬂ «·”Ì«—…"
         Height          =   285
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label xGas_differ 
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label27 
         Caption         =   "›—ﬁ «·«” Â·«ﬂ"
         Height          =   285
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   675
         Width           =   1680
      End
      Begin VB.Label xGasPerKilo 
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
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label25 
         Caption         =   "„⁄œ· «” Â·«ﬂ «·”Ì«—…"
         Height          =   285
         Left            =   6885
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1095
      Left            =   16110
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   6570
      Width           =   2850
      Begin VB.TextBox xDate_Policy 
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
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "D"
         Top             =   585
         Width           =   1545
      End
      Begin VB.TextBox xPolicy 
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
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label18 
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label16 
         Caption         =   "—ﬁ„ «·»Ê·Ì’…"
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   270
         Width           =   1110
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid3 
      Height          =   1995
      Left            =   90
      TabIndex        =   92
      Top             =   1080
      Width           =   6675
      _cx             =   11774
      _cy             =   3519
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
      RowHeightMin    =   300
      RowHeightMax    =   300
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   270
      Top             =   9180
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
      Caption         =   "DATA5"
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
   Begin VB.Label Label30 
      Caption         =   "„·ÕÊŸ… "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11700
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   90
      Width           =   795
   End
End
Attribute VB_Name = "Travelfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, bSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String
Dim cFilter As String, cList As String, clist1 As String, cList2 As String, cPlaceString As String
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchClient As New Search3, oSearchCar As New Search3, oSearchSup As New Search3, oSearchBox As New Search3, oSearchDriver As New Search3
Dim bedit As Boolean
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CODE", addstring(xCode.Text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xdesca.Text))
aInsert = AddFlag(aInsert, "[CAR]", addstring(xcar.Text))
aInsert = AddFlag(aInsert, "[DISTANCE]", Val(xDistance.Text))
aInsert = AddFlag(aInsert, "[TOTAL]", Val(xTotal.Text))
aInsert = AddFlag(aInsert, "[TRAILER]", addvalue(xTrailer.BoundText))
aInsert = AddFlag(aInsert, "[PLACE1]", addvalue(xPlace1.BoundText))
aInsert = AddFlag(aInsert, "[PLACE2]", addvalue(xPlace2.BoundText))
aInsert = AddFlag(aInsert, "[POLICY]", addstring(xPolicy.Text))
aInsert = AddFlag(aInsert, "[DATE_POLICY]", addDate(xDate_Policy.Text))
aInsert = AddFlag(aInsert, "[DRIVER]", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "[DRIVER2]", addstring(xDriver2.BoundText))
aInsert = AddFlag(aInsert, "[FOLLOWER]", addstring(xFollower.BoundText))
aInsert = AddFlag(aInsert, "[CARGO]", addvalue(xCargo.BoundText))
aInsert = AddFlag(aInsert, "[WEIGHT]", Val(xWeight.Text))
'aInsert = AddFlag(aInsert, "[CLASS]", IIf(Val(xTotal_weight.Caption) <> 0, 1, Val(xClass.Text)))
aInsert = AddFlag(aInsert, "[CLASS]", Val(xClass.Text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.Text))
aInsert = AddFlag(aInsert, "[TRUST_CLOSE]", xTrust_close.Value)
aInsert = AddFlag(aInsert, "[CODE_SUP]", addstring(xCode_sup.Text))
aInsert = AddFlag(aInsert, "[TOTAL_SUP]", Val(xTotal_sup.Text))
aInsert = AddFlag(aInsert, "[GAS]", Val(xGas.Text))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Text = "" Then
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(sUserName))
    xDoc_No.Text = RetZero(Val(Newflag("TRAVEL_H", "doc_no")))
    aInsert = AddFlag(aInsert, "[DOC_NO]", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "TRAVEL_H")
Else
    con.Execute addUpdate(aInsert, "TRAVEL_H", "doc_no = " & addstring(xDoc_No.Text))
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
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
        Unload oSearchBox
        Grid1_AfterEdit grid1.Row, grid1.Col
        CellPos 13, grid1.Row, 0
    ElseIf grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 1)
        Unload oSearchItem
        Grid1_AfterEdit grid1.Row, grid1.Col
        CellPos 13, grid1.Row, 1
    ElseIf grid1.Col = 6 Then
        grid1.TextMatrix(grid1.Row, 6) = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 0) = ""
        Unload oSearchSup
        Grid1_AfterEdit grid1.Row, grid1.Col
        CellPos 13, grid1.Row, 6
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
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    xcode_Desca.Caption = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 1)
    Unload oSearchClient
    SendKeys "{TAB}"
ElseIf ActiveControl.Name = xCode_sup.Name Then
    xCode_sup.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    xCode_sup_desca.Caption = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 1)
    Unload oSearchSup
    xTotal_sup.SetFocus
ElseIf ActiveControl.Name = xcar.Name Then
    xcar.Text = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 0)
    xCar_Validate False
    Unload oSearchCar
    SendKeys "{TAB}"
ElseIf ActiveControl.Name = xDriver.Name Then
    xDriver.BoundText = oSearchDriver.grid1.TextMatrix(oSearchDriver.grid1.Row, 0)
    Unload oSearchDriver
    SendKeys "{TAB}"
ElseIf ActiveControl.Name = xFollower.Name Then
    xFollower.BoundText = oSearchDriver.grid1.TextMatrix(oSearchDriver.grid1.Row, 0)
    Unload oSearchDriver
    SendKeys "{TAB}"
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

Private Sub cmdCalcTotal_Click()
If Not (xPlace1.MatchedWithList And xPlace2.MatchedWithList) Then
    MsgBox "„ﬂ«‰ «·ﬁÌ«„ Ê«·Ê’Ê· €Ì— „”Ã·Ì‰"
    Exit Sub
End If
travel_weightfrm.sDoc_no = xDoc_No.Text
travel_weightfrm.Show 1
openCardTable
myUndo
End Sub

Private Sub cmdCargo_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sCode = xCargo.BoundText
oFlagfrm.sTable = "CARGO_CODES"
oFlagfrm.sCaption = "«‰Ê«⁄ «·Õ„Ê·« "
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA3.Recordset = myRecordSet("SELECT * FROM CARGO_CODES", con)
xCargo.BoundText = sCode
If Not xCargo.MatchedWithList Then xCargo.BoundText = ""
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From TRAVEL_C where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRAVEL_T where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRAVEL_H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
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
Travel_LookupAll Me, oSearchDoc, cFilter
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

Private Sub cmdPlace_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String, sCode2 As String
sCode = xPlace1.BoundText
sCode2 = xPlace2.BoundText
oFlagfrm.sTable = "PLACE_CODES"
oFlagfrm.sCaption = "«·«„«ﬂ‰"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set data4.Recordset = myRecordSet("SELECT * FROM PLACE_CODES ORDER BY DESCA", con)
xPlace1.BoundText = sCode
xPlace2.BoundText = sCode2
If Not xPlace1.MatchedWithList Then xPlace1.BoundText = ""
If Not xPlace2.MatchedWithList Then xPlace2.BoundText = ""
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
        cString = cString & turn(cString) & " doc_no = " & MyParn(loctable!DOC_NO)
        con.Execute cString
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing


Set loctable = New ADODB.Recordset
loctable.Open "Select DOC_NO,SUM(PRICE * QUANT) AS TOTAL FROM TRAVEL GROUP BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    con.Execute "UPDATE TRAVEL_H SET TRAVEL_H.CASH = " & Val(loctable!TOTAL & "") & " WHERE DOC_NO = " & MyParn(loctable!DOC_NO)
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
    Grid1_Validate False
    cmdSave_Click
    KeyCode = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Name <> xPolicy.Name Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con
cList = StrList2("Select code,desca from FILE0_50  order by desca")
clist1 = StrList("Select code,desca from FILE0_50 WHERE CODE >= '500001'order by desca")
cList2 = StrList("Select code,desca from FILE0_50 WHERE CODE < '500001'order by desca")

Fixgrd3
'cList2 = StrList2("Select code,desca from FILE4_10 order by desca")

'DATA1.ConnectionString = strCon
'DATA1.RecordSource = "SELECT * FROM DRIVER WHERE DRIVER = 1 ORDER BY DESCA"
Set data1.Recordset = myRecordSet("SELECT * FROM DRIVER WHERE DRIVER = 1 ORDER BY DESCA", con)
Set xDriver.RowSource = data1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

Set xDriver2.RowSource = data1
xDriver2.ListField = "Desca"
xDriver2.BoundColumn = "Code"

'DATA2.ConnectionString = strCon
'DATA2.RecordSource = "SELECT * FROM DRIVER ORDER BY DESCA"
Set DATA2.Recordset = myRecordSet("SELECT * FROM DRIVER ORDER BY DESCA", con)
Set xFollower.RowSource = DATA2
xFollower.ListField = "Desca"
xFollower.BoundColumn = "Code"

'DATA3.ConnectionString = strCon
'DATA3.RecordSource = "SELECT * FROM CARGO_CODES"
Set DATA3.Recordset = myRecordSet("SELECT * FROM CARGO_CODES", con)
Set xCargo.RowSource = DATA3
xCargo.ListField = "Desca"
xCargo.BoundColumn = "Code"

'cPlaceString = "SELECT DISTINCT LTRIM(RTRIM(PLACE1)) AS PLACE " & _
'          "  From dbo.TRAVEL_H" & _
'          "  Where (Not (PLACE1 Is Null))" & _
'          "  Union" & _
'          "  SELECT DISTINCT LTRIM(RTRIM(PLACE2)) AS PLACE " & _
'          "  FROM         dbo.TRAVEL_H AS TRAVEL_H_1" & _
'          "  Where (Not (PLACE2 Is Null))"


Set DATA5.Recordset = myRecordSet("SELECT * FROM TRAILER_CODES ORDER BY DESCA", con)
Set xTrailer.RowSource = DATA5
xTrailer.ListField = "DESCA"
xTrailer.BoundColumn = "CODE"

Set data4.Recordset = myRecordSet("SELECT * FROM PLACE_CODES ORDER BY DESCA", con)
Set xPlace1.RowSource = data4
xPlace1.ListField = "DESCA"
xPlace1.BoundColumn = "CODE"

Set xPlace2.RowSource = data4
xPlace2.ListField = "DESCA"
xPlace2.BoundColumn = "CODE"

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

Set grid2.DataSource = DATA12
DATA12.ConnectionString = strCon

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
Unload oSearchClient

Set Travelfrm = Nothing
Err.Clear
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bedit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM TRAVEL_C WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
        Calctotals
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 112 And bedit = True And grid1.Col = 0 Then
    BoxLookupAll Me, oSearchBox, "FILE0_50.CODE < '500001'"
ElseIf KeyCode = 112 And bedit = True And grid1.Col = 1 Then
    ChargeLookupAll Me, oSearchItem
ElseIf KeyCode = 112 And bedit = True And grid1.Col = 6 Then
    SupLookupAll Me, oSearchSup
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
If Row = .Rows - 1 Then myAddItem

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
    grid1.TextMatrix(Row, 2) = GetField("SELECT DESCA FROM FILE8_51 WHERE CODE = " & MyParn(grid1.TextMatrix(Row, 1))) & ""
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
If Trim(.TextMatrix(Row, 1)) = "" And Trim(.TextMatrix(Row, 6)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
'If Not IsDate(.TextMatrix(Row, 3)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 4)) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        grid1.RemoveItem OldRow
        'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
        Calctotals
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
If xClosed.Value = 1 Or bedit = False Then
    grid1.Editable = flexEDNone
    Exit Sub
ElseIf .Col = 0 Or .Col = 1 Or .Col = 4 Or .Col = 5 Then
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
Private Sub Grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
    'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
    Calctotals
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        .EditText = RetZero(.EditText, 3)
        Dim aRet As Variant
        aRet = GetFields("Select code,desca from file8_51 where code = " & MyParn(.EditText))
        If IsEmpty(aRet) Then
            MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
        Else
            .TextMatrix(Row, 2) = retFlag(aRet, "DESCA") & ""
        End If
    End If
ElseIf Col = 0 Then
    If Trim(.EditText) <> "" And Trim(.TextMatrix(Row, 6)) <> "" Then
        .TextMatrix(Row, 6) = ""
    End If
ElseIf Col = 3 Then
    If Not IsDate(.EditText) Then
        Cancel = True
    Else
        .EditText = Format(.EditText, "YYYY/MM/DD")
    End If
ElseIf Col = 6 Then
    If Trim(.EditText) <> "" And Trim(.TextMatrix(Row, 0)) <> "" Then
        .TextMatrix(Row, 0) = ""
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·Œ“‰…|" & "«·ﬂÊœ|" & "«·„’—Ê›|" & "«· «—ÌŒ|" & "«·ﬁÌ„…|" & "«·»Ì«‰|" & "«·„Ê—œ|"
.ColWidth(0) = 2000
.ColWidth(1) = 800
.ColWidth(2) = 2000
.ColWidth(3) = 1400
.ColWidth(4) = 900
.ColWidth(5) = 2800
.ColWidth(6) = 2000
.ColWidth(7) = 300
.ColComboList(0) = cList
.ColHidden(3) = True
.ColHidden(6) = True
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Fixgrd3()
With grid3
.FormatString = "«·ﬂÊœ|" & "«·«”„|" & "«·⁄Âœ…|" & "«·„’—Ê›|" & "«·›—ﬁ"
.ColWidth(0) = 2000
.ColWidth(1) = 3000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColComboList(1) = cList
.ColHidden(0) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub

Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 3 Then
    .Col = Col + 1 + IIf(Col = 1, 2, 0) + IIf(Col = 2, 1, 0)
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, IIf(.TextMatrix(Row + 1, 0) <> "", 1, 0)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
.TextMatrix(.Rows - 1, 0) = xDriver.BoundText
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "BOX", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "CHARGE", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
        aInsert = AddFlag(aInsert, "[VALUE]", Val(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(i, 6)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "TRAVEL_C")
        Else
            con.Execute addUpdate(aInsert, "TRAVEL_C", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
cString = "SELECT TRAVEL_C.BOX,TRAVEL_C.CHARGE,FILE8_51.DESCA, CONVERT(VARCHAR(10),[DATE],111),TRAVEL_C.[VALUE],TRAVEL_C.DESCA,TRAVEL_C.CODE,TRAVEL_C.ID " & _
          " FROM TRAVEL_C INNER JOIN FILE8_51 ON TRAVEL_C.CHARGE = FILE8_51.CODE"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.Text)
DATA11.RecordSource = cString
DATA11.Refresh
myAddItem
End With
Calctotals
Fixgrd
End Sub

Private Sub xCar_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then carsLookupAll Me, oSearchCar
End Sub

Private Sub xCar_Validate(Cancel As Boolean)
xCar_Desca.Caption = ""
xCar_type.Caption = ""
xGas1.Caption = ""
If Not ValidInt(xcar.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select code,[TYPE],MODEL,BOARD,gas1 from cars where code = " & xcar.Text)
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·”Ì«—… €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xCar_Desca.Caption = retFlag(aRet, "Board")
    xCar_type.Caption = retFlag(aRet, "TYPE") & " " & retFlag(aRet, "Model")
    xGas1.Caption = retFlag(aRet, "gas1")
End If
Calctotals2
End Sub
Private Sub xClass_LostFocus()
myLostFocus xClass
Calctotals
End Sub
Private Sub xClass_Validate(Cancel As Boolean)
'If Me.xTotal_weight.Value <> 0 And Val(xClass.Text) <> 1 Then
'    MsgBox "ﬁÌ„… «·›∆… ÌÃ» «‰  ﬂÊ‰ 1"
'    Cancel = True
'End If
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookupAll Me, oSearchClient
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub

Private Sub xCode_sup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub

Private Sub xCode_sup_Validate(Cancel As Boolean)
xCode_sup_desca.Caption = ""
If xCode_sup.Text = "" Then Exit Sub
xCode_sup.Text = RetZero(xCode_sup.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file4_10 where code = " & MyParn(xCode_sup.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xCode_sup_desca.Caption = retFlag(aRet, "desca") & ""
End If
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
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
Dim i As Integer
If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Trim(xCode.Text) = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If
If Val(xTotal.Text) < 0 Then
    If Not bIgMsg Then MsgBox "«·‰Ê·Ê‰ «ﬁ· „‰ ’›—"
    Exit Function
End If

If IsDate(xDate_Policy.Text) And Trim(xPolicy.Text) = "" Then
    If Not bIgMsg Then MsgBox "Â‰«ﬂ  «—ÌŒ »Ê·Ì’… Ê·« ÌÊÃœ —ﬁ„ »Ê·Ì’… ‘Õ‰"
    Exit Function
End If

'If (Not IsDate(xDate_Policy.Text)) And Trim(xPolicy.Text) <> "" Then
'    If Not bIgMsg Then MsgBox "Â‰«ﬂ —ﬁ„ »Ê·Ì’… ‘Õ‰ Ê·« ÌÊÃœ  «—ÌŒ ‘Õ‰"
'    Exit Function
'End If

With grid1
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!DOC_NO
xdesca.Text = CardTable!desca & ""
xPolicy.Text = CardTable!Policy & ""
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xDate_Policy.Text = Format(CardTable!Date_policy, "YYYY-MM-DD")
xDriver.BoundText = CardTable!Driver & ""
xDriver2.BoundText = CardTable!Driver2 & ""
xTrailer.BoundText = CardTable!trailer & ""
xNotes.Text = CardTable!Notes & ""
xCode.Text = CardTable!code & ""
xcode_Desca.Caption = CardTable!code_Desca & ""
xcar.Text = CardTable!car & ""
xFollower.BoundText = CardTable!follower & ""
xWeight_Total.Caption = Myvalue(CardTable!weight_total)

xWeight.Text = Myvalue(CardTable!Weight)
xClass.Text = Myvalue(CardTable!Class)

xCargo.BoundText = CardTable!cargo & ""
xPlace1.BoundText = CardTable!place1 & ""
xPlace2.BoundText = CardTable!place2 & ""
xDistance.Text = Myvalue(CardTable!Distance)
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
xTrust_close.Value = IIf(CardTable!Trust_Close, 1, 0)
xCode_sup.Text = CardTable!CODE_SUP & ""
xCode_sup_Validate False
xTotal_sup.Text = Myvalue(CardTable!TOTAL_SUP)
xGas.Text = Myvalue(CardTable!gas)
xWeight_Value.Caption = Myvalue(CardTable!weight_value)
xDiscount.Caption = Myvalue(CardTable!Discount)
xExtend.Caption = Myvalue(CardTable!Extend)
xTotal_weight.Caption = Myvalue(CardTable!weight_value + CardTable!Extend - CardTable!Discount)
xCar_Validate False
Calctotals2
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
myloadgrd2
cellpos2 13, grid2.Rows - 2, grid2.Cols - 1
Handlecontrols LoadMode
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Sub mydefine()
xDoc_No.Text = ""
xdesca.Text = ""
xPolicy.Text = ""
xDate.Text = ""
xDate_Policy.Text = ""
xDriver.BoundText = ""
xDriver2.BoundText = ""
xNotes.Text = ""
xCode.Text = ""
xcode_Desca.Caption = ""
xCar_type.Caption = ""
xcar.Text = ""
xCar_Desca.Caption = ""
xPlace1.BoundText = ""
xPlace2.BoundText = ""
xDistance.Text = ""
xWeight.Text = ""
xClass.Text = ""
xFollower.BoundText = ""
xTrailer.BoundText = ""
xCargo.BoundText = ""
xTotal.Text = ""
xTotal_Cost.Caption = ""
xProfit.Caption = ""
xProfit_rate.Caption = ""
xTotal_trust.Caption = ""
xRest_trust.Caption = ""
xCode_sup.Text = ""
xCode_sup_desca.Caption = ""
xTotal_sup.Text = ""
xTrust_close.Value = 0
xGas.Text = ""
xClosed.Value = 0
grid1.Rows = 1
myAddItem
Fixgrd

grid2.Rows = 1
MyAddItem2
fixgrd2

grid3.Rows = 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdAdd.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDel.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
CmdInform.Enabled = sDoc_no = ""
cmdCalcTotal.Enabled = nMode = LoadMode
'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xdesca
End Sub

Private Sub xDesca_LostFocus()
myLostFocus xdesca
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
Dim nTotal_Cost As Double, nTotal_Trust As Single, nReal_cost As Double
xTotal.Text = Myvalue(Val(xClass.Text) * Val(xWeight.Text))
With grid1
grid3.Rows = 1
For i = 1 To .Rows - 2
    If Trim(.TextMatrix(i, 0)) <> "" Then
        nFound = grid3.FindRow(.TextMatrix(i, 0), , 0)
        If nFound = -1 Then
            grid3.AddItem ""
            grid3.TextMatrix(grid3.Rows - 1, 0) = .TextMatrix(i, 0)
            grid3.TextMatrix(grid3.Rows - 1, 1) = .TextMatrix(i, 0)
            grid3.TextMatrix(grid3.Rows - 1, 3) = Val(.TextMatrix(i, 4))
        Else
            grid3.TextMatrix(nFound, 3) = Val(grid3.TextMatrix(nFound, 3)) + Val(.TextMatrix(i, 4))
        End If
    End If
    nTotal_Cost = Val(.TextMatrix(i, 4)) + nTotal_Cost
Next
xTotal_Cost.Caption = Myvalue(nTotal_Cost)
End With

With grid2
For i = 1 To .Rows - 2
    If Trim(.TextMatrix(i, 2)) <> "" Then
        nFound = grid3.FindRow(.TextMatrix(i, 2), , 0)
        If nFound = -1 Then
            grid3.AddItem ""
            grid3.TextMatrix(grid3.Rows - 1, 0) = .TextMatrix(i, 2)
            grid3.TextMatrix(grid3.Rows - 1, 1) = .TextMatrix(i, 2)
            grid3.TextMatrix(grid3.Rows - 1, 2) = Val(.TextMatrix(i, 4))
        Else
            grid3.TextMatrix(nFound, 2) = Val(grid3.TextMatrix(nFound, 2)) + Val(.TextMatrix(i, 4))
        End If
    End If
    nTotal_Trust = Val(.TextMatrix(i, 4)) + nTotal_Trust
Next

xTotal_trust.Caption = Myvalue(nTotal_Trust)
'If nTotal_Trust <> 0 Then
    xRest_trust.Caption = Round(nTotal_Trust - nTotal_Cost, 2)
'Else
'    xRest_trust.Caption = ""
'End If
End With

nReal_cost = IIf(Trim(xCode_sup.Text) = "", Val(xTotal_Cost.Caption), Val(xTotal_sup.Text))
xProfit.Caption = Round(Val(xTotal.Text) - nReal_cost, 2)

If Val(xTotal.Text) <> 0 Then
    xProfit_rate.Caption = Myvalue(Round(Val(xProfit.Caption) / Val(xTotal.Text) * 100, 2))
Else
    xProfit_rate.Caption = ""
End If

With grid3
For i = 1 To .Rows - 1
    .TextMatrix(i, 4) = Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 3))
Next
If .Rows > 1 Then
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#0.00", &HE0E0E0, , True, "  "
    .Subtotal flexSTSum, -1, 3, "#0.00", &HE0E0E0, , True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", &HE0E0E0, , True, "  "
    .TextMatrix(.Rows - 1, 1) = "«·≈Ã„«·Ï"
End If
End With
'If Not IsEmpty(aCost) Then
'    For I = 0 To UBound(aCost) Step 2
'        If grid3.FindRow(aCost(I), , 0) Then
'            grid3.TextMatrix(
'        End If
'    Next
'End If
End Function
Private Function Calctotals2()
If Val(xDistance.Text) <> 0 Then
    xGasPerKilo.Caption = Myvalue(Val(xGas.Text) / Val(xDistance.Text))
Else
    xGasPerKilo.Caption = ""
End If
xGasPerKilo_Differ.Caption = Myvalue((Val(xGasPerKilo.Caption) - Val(xGas1.Caption)) * 100)
xGascar.Caption = Myvalue(Val(xDistance.Text) * Val(xGas1.Caption))
xGas_differ.Caption = Myvalue(Val(xGas.Text) - Val(xGascar.Caption))
End Function

Private Sub CardLookup(Optional pWhere As String = "")
End Sub
Private Function mysave() As Boolean
If Not MYVALID Then Exit Function
Calctotals
If Not myreplace Then Exit Function
Inform " „ Õ›Ÿ «·„” ‰œ"
If sDoc_no <> "" Then
    Unload Me
Else
    openCardTable
    myUndo
End If
End Function
Private Function doprint() As Boolean
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
cmdSave.Enabled = (bedit)
'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TRAVEL_H.*,FILE3_10.DESCA AS CODE_DESCA,CARS.BOARD,CARS.TYPE,CARS.MODEL FROM TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.Code = FILE3_10.CODE LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE"
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

Private Sub xDriver_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then DriverLookupAll Me, oSearchDriver, "DRIVER.DRIVER = 1"
End Sub
Private Sub xDriver_LostFocus()
myLostFocus xDriver
If Not xDriver.MatchedWithList Then
    xDriver.BoundText = RetZero(xDriver.Text)
    If Not xDriver.MatchedWithList Then xDriver.BoundText = ""
End If

If xDriver.MatchedWithList Then
    grid1.TextMatrix(grid1.Rows - 1, 0) = xDriver.BoundText
    grid2.TextMatrix(grid2.Rows - 1, 2) = xDriver.BoundText
End If
End Sub
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With grid2
    If KeyCode = 46 And .Row <> .Rows - 1 And bedit Then
        If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.BeginTrans
                On Error GoTo myerror
                con.Execute "DELETE FROM TRAVEL_T WHERE ID = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            .RemoveItem .Row
            Calctotals
        End If
    ElseIf KeyCode = 13 Then
        cellpos2 KeyCode, .Row, .Col
    ElseIf KeyCode = 112 And bedit = True And (.Col = 0) Then
        BoxLookupAll Me, oSearchBox, "FILE0_50.CODE >= '500001'"
    ElseIf KeyCode = 112 Then
        BoxLookupAll Me, oSearchBox, "FILE0_50.CODE < '500001'"
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
        If .TextMatrix(.Rows - 2, 2) <> "" Then
            .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 2, 2)
        End If
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
If KeyCode = 13 And Col <> 2 Then
    cellpos2 KeyCode, Row, Col
End If
End Sub
Private Function validRow2(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 3)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 4)) Then Exit Function
End With
validRow2 = True
End Function
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid2.Rows - 1 And OldRow <> 0 And grid2.TextMatrix(OldRow, grid2.Cols - 1) = "" Then
    If Not validRow2(OldRow) Then
        grid2.RemoveItem OldRow
        'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
        Calctotals
    End If
End If
End Sub
Private Sub grid2_EnterCell()
With grid2
If xClosed.Value = 1 Or bedit = False Then
    .Editable = flexEDNone
    Exit Sub
End If
If (.Col = 0 Or .Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5) Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub grid2_GotFocus()
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
    grid2.RemoveItem grid2.Row
    'xDriver.Enabled = grid1.Rows = 2 And grid2.Rows = 2
    Calctotals
End If
End Sub
Private Sub GRID2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim aRet As Variant
With grid2
If Col = 0 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        .EditText = RetZero(.EditText, 6)
        aRet = GetFields("Select code,desca from file0_50 where code = " & MyParn(.EditText))
        If IsEmpty(aRet) Then
            MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
        Else
            .TextMatrix(Row, 1) = retFlag(aRet, "DESCA") & ""
        End If
    End If
ElseIf Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 2 Then
    aRet = GetFields("Select code,desca from file0_50 where code = " & MyParn(RetZero(.EditText)))
    If Not IsEmpty(aRet) Then
        .EditText = RetZero(.EditText)
    End If
ElseIf Col = 3 Then
    If Not IsDate(.EditText) Then
        Cancel = True
    Else
        .EditText = Format(.EditText, "YYYY/MM/DD")
    End If
ElseIf Col = 4 Then
    If Not IsNumeric(.EditText) Then
        MsgBox "ﬁÌ„… €Ì— „”Ã·…"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub fixgrd2()
With grid2
.FormatString = "„‰ Œ“‰…|" & "„‰ Œ“‰…|" & "≈·Ì Œ“‰…|" & "«· «—ÌŒ|" & " «·„»·€|" & "«·»Ì«‰|"
.ColWidth(0) = 900
.ColWidth(1) = 2000
.ColWidth(2) = 2000
.ColWidth(3) = 1400
.ColWidth(4) = 800
.ColWidth(5) = 2000
.ColComboList(2) = cList2
.MergeCells = flexMergeFixedOnly
.MergeRow(0) = True
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub cellpos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid2
KeyCode = 0
If Col < .Cols - 2 Then
    If Col < 3 Then
        .Col = NextEmpty(grid2, Row, Col + 1, 3)
    Else
        .Col = Col + 1
    End If
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid2, Row + 1, 0, 3)
    .ShowCell Row + 1, 1
ElseIf Row = grid2.Rows - 1 And Col >= grid2.Cols - 2 Then
    grid1.SetFocus
End If
End With
End Sub
Private Sub MyAddItem2()
With grid2
.AddItem ""
If .Rows > 2 Then
    If .TextMatrix(.Rows - 2, 0) <> "" Then .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
    If .TextMatrix(.Rows - 2, 1) <> "" Then .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
    If .TextMatrix(.Rows - 2, 2) <> "" And Not xDriver.MatchedWithList Then .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 2, 2)
End If
If xDriver.MatchedWithList Then .TextMatrix(.Rows - 1, 2) = xDriver.BoundText
End With
End Sub
Private Function myreplaceGrd2(Row) As Boolean
Dim aInsert As Variant
With grid2
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "[BOX1]", addstring(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[BOX2]", addstring(.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "[VALUE]", Val(.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(.TextMatrix(i, 5)))
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "TRAVEL_T")
        Else
            con.Execute addUpdate(aInsert, "TRAVEL_T", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd2 = True
End Function
Private Sub myloadgrd2()
With grid2
cString = "SELECT TRAVEL_T.BOX1,FILE0_50.DESCA,TRAVEL_T.BOX2,CONVERT(VARCHAR(10),TRAVEL_T.DATE,111),TRAVEL_T.VALUE,TRAVEL_T.[DESCA],ID " & _
          " FROM TRAVEL_T LEFT JOIN FILE0_50 ON TRAVEL_T.BOX1 = FILE0_50.CODE"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
DATA12.RecordSource = cString
DATA12.Refresh
MyAddItem2
End With
Calctotals
fixgrd2
End Sub

Private Sub xFollower_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then DriverLookupAll Me, oSearchDriver, , "LAST_CLICKED"
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_Policy_GotFocus()
myGotFocus xDate_Policy
End Sub
Private Sub xDate_Policy_LostFocus()
myLostFocus xDate_Policy
myValidDate xDate_Policy
End Sub
Private Sub xPolicy_GotFocus()
myGotFocus xPolicy
End Sub
Private Sub xPolicy_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 '   On Error Resume Next
    KeyCode = 0
    xDate_Policy.SetFocus
    Err.Clear
End If
End Sub

Private Sub xPolicy_LostFocus()
myLostFocus xPolicy
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
myGotFocus xcar
End Sub
Private Sub xCar_LostFocus()
myLostFocus xcar
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

Private Sub xPlace1_GotFocus()
myGotFocus xPlace1
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
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

Private Sub xTotal_net_Click()

End Sub

Private Sub xTrailer_GotFocus()
myGotFocus xTrailer
End Sub
Private Sub Xweight_GotFocus()
myGotFocus xWeight
End Sub
Private Sub Xweight_LostFocus()
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
Private Sub xTrailer_LostFocus()
myLostFocus xTrailer
If Not xTrailer.MatchedWithList Then xTrailer.BoundText = ""
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
Private Sub xPlace2_LostFocus()
myLostFocus xPlace2
If Not xPlace2.MatchedWithList Then
    If Trim(xPlace2.Text) = xPlace2.BoundText Then
        xPlace2.BoundText = xPlace2.Text
    Else
        xPlace2.BoundText = ""
    End If
End If
End Sub
Private Sub xPlace1_LostFocus()
myLostFocus xPlace1
If Not xPlace1.MatchedWithList Then
    If Trim(xPlace1.Text) = xPlace1.BoundText Then
        xPlace1.BoundText = xPlace1.Text
    Else
        xPlace1.BoundText = ""
    End If
End If
End Sub

