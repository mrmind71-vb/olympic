VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form itemsfrm8 
   Caption         =   "»‰Êœ «·«‘ —«ﬂ« "
   ClientHeight    =   9045
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   11880
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
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   3015
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   45
      Width           =   1545
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   45
         Picture         =   "items8.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      Height          =   690
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   45
      Width           =   7215
      Begin VB.CommandButton cmdsave 
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
         Picture         =   "items8.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
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
         Picture         =   "items8.frx":478D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Picture         =   "items8.frx":6D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "items8.frx":9172
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "items8.frx":BA0C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "items8.frx":DFB8
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   3765
      Left            =   360
      TabIndex        =   26
      Top             =   3960
      Width           =   11400
      _cx             =   20108
      _cy             =   6641
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
      FixedRows       =   2
      FixedCols       =   2
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
      AutoResize      =   -1  'True
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame5 
      Caption         =   "»‰œ Œ«’ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CheckBox xShowPaid 
         Appearance      =   0  'Flat
         Caption         =   " ŸÂ— ›Ï «·„ÿ«·»…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   315
         Width           =   1830
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "€—«„« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2880
      Width           =   3030
      Begin VB.TextBox xDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   495
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   585
         Width           =   360
      End
      Begin VB.CheckBox xMeeting 
         Appearance      =   0  'Flat
         Caption         =   "«·»‰œ €—«„… Ã„⁄Ì… ⁄„Ê„Ì…"
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
         Height          =   240
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   315
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÌÊ„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   6
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "€—«„…  √ŒÌ— ”œ«œ »⁄œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   630
         Width           =   1935
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1395
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   405
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "‘"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "«·ﬁ—«»…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2160
      Width           =   5640
      Begin VB.TextBox xAge2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   3825
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox xAge1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   3825
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   990
         Width           =   555
      End
      Begin MSDBCtls.DBCombo xRelation 
         Bindings        =   "items8.frx":1078B
         Height          =   315
         Left            =   1575
         TabIndex        =   3
         Top             =   270
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   1575
         TabIndex        =   41
         Top             =   630
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label12 
         Caption         =   "«·‰Ê⁄"
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·‰ ”‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   14
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ ”‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   13
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1395
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»‰œ «·ﬁ—«»…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   5640
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   90
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   4785
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4320
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   4320
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   5
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   435
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   630
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame6 
      Caption         =   "ÿ»Ì⁄… «·»‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1710
      Width           =   2670
      Begin VB.CheckBox xBasicOld 
         Appearance      =   0  'Flat
         Caption         =   "«”«”Ì ··⁄÷Ê «·ﬁœÌ„ ›ﬁÿ"
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
         Height          =   225
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1845
         Width           =   2235
      End
      Begin VB.CheckBox xNoRate 
         Appearance      =   0  'Flat
         Caption         =   "·« ÌÕ”» ⁄·ÌÂ ”‰… Ê‰’"
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   2205
      End
      Begin VB.CheckBox xBasicDied 
         Appearance      =   0  'Flat
         Caption         =   "√”«”Ì ·√»‰«¡ «·„ Ê›Ì "
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1500
         Width           =   2010
      End
      Begin VB.CheckBox xBasicNew 
         Appearance      =   0  'Flat
         Caption         =   "√”«”Ì ··⁄÷Ê «·ÃœÌœ"
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox xLate 
         Appearance      =   0  'Flat
         Caption         =   "⁄·ÌÂ €—«„…"
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   1290
      End
      Begin VB.CheckBox xAllMember 
         Appearance      =   0  'Flat
         Caption         =   "··ﬂ· "
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
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   900
         Width           =   705
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   7740
      Width           =   4290
      Begin VB.CommandButton CmdLast 
         Caption         =   "√ŒÌ—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   840
      End
      Begin VB.CommandButton CmdFirst 
         Caption         =   "√Ê·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   975
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "·«Õﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2325
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton CmdPrevious 
         Caption         =   "”«»ﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3300
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   135
         Width           =   915
      End
   End
   Begin Threed.SSCommand cmdBranch 
      Height          =   375
      Left            =   10440
      TabIndex        =   45
      Top             =   7785
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      _Version        =   196610
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«Œ Ì«— «·›—⁄"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdNoBranch 
      Height          =   375
      Left            =   7380
      TabIndex        =   46
      Top             =   8190
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
      _Version        =   196610
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«·€«¡ «·«Œ Ì«—"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "itemsfrm8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As Recordset, Deletetable1 As Recordset
Dim oSearch As New Search3
Dim SectionItemTable As Recordset
Dim RecordCountTable As Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select code,[desca] from file2_10"
Generalarray(2) = "Order by FILE2_10.CODE"
Generalarray(3) = 7000
Generalarray(5) = True

listarray(0, 0) = "«·ﬂÊœ-«·»Ì«‰"
listarray(0, 1) = "(VAL('cFilter') = CODE OR %%DESCA%%)"


GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.bEnter = False
oSearch.Caption = "«” ⁄·«„ «·»‰Êœ"
oSearch.Show 1
End Sub
Sub myDefine()
xDescA.Text = ""
xValue.Text = ""

xIsMember.Value = 0
xRelation.BoundText = ""
xAge1.Text = ""
xAge2.Text = ""
xSex.ListIndex = 0

xNoRate.Value = 0
xLate.Value = 0
xAllMember.Value = 0
xBasicNew.Value = 0
xBasicDied.Value = 0
xShowPaid.Value = 0

xLocker.BoundText = ""
xYacht.BoundText = ""

xMeeting.Value = 1
xDays.Text = ""

xBasicOld.Value = 0


With grid1
For i = 1 To grid1.rows - 1
    .TextMatrix(i, 2) = ""
    .TextMatrix(i, 3) = ""
     For nCol = 4 To grid1.Cols - 1
         .TextMatrix(i, nCol) = ""
     Next
Next
End With

Handlecontrols DefineMode
If myRecordCount = 0 Then
    xRecordNumber = ""
Else
    xRecordNumber = "”Ã· " & myRecordCount + 1 & " „‰ " & myRecordCount + 1
End If
End Sub
Sub myProc()
CardTable.FindFirst "Code = " & oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
myload
Unload oSearch
End Sub
Sub myload()
xCode.Text = CardTable.CODE
xDescA.Text = TurnValue(CardTable.DESCA, Null, "")
xValue.Text = Val(CardTable!Value & "")
xIsMember.Value = IIf(CardTable!isMember, 1, 0)
xShowPaid.Value = IIf(CardTable!showpaid, 1, 0)

xRelation.BoundText = TurnValue(CardTable!Relation, Null, "")

xAge1.Text = TurnValue(CardTable!age1, Null, "")
xAge2.Text = TurnValue(CardTable!age2, Null, "")
xDays.Text = TurnValue(CardTable!days, Null, "")

xSex.ListIndex = LoadSex(CardTable!Sex)

xNoRate.Value = IIf(CardTable!NoRate, 1, 0)
xLate.Value = IIf(CardTable!LATE, 1, 0)
xAllMember.Value = IIf(CardTable!AllMember, 1, 0)
xBasicNew.Value = IIf(CardTable!BasicNew, 1, 0)
xBasicDied.Value = IIf(CardTable!BasicDied, 1, 0)
xBasicOld.Value = IIf(CardTable!BASICOLD, 1, 0)
xLocker.BoundText = TurnValue(CardTable!Locker, Null, "")
xYacht.BoundText = TurnValue(CardTable!yacht, Null, "")
xMeeting.Value = IIf(CardTable!Meeting, 1, 0)
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & myRecordCount

LoadGrd
Handlecontrols LoadMode
End Sub
Sub myreplace()
Me.MousePointer = 11

If xRelation.BoundText <> "" Or xLocker.BoundText <> "" Or xYacht.BoundText <> "" Then
    xAllMember.Value = 0
End If

xRelation.BoundText = IIf(xLocker.BoundText <> "" Or xYacht.BoundText <> "", "", xRelation.BoundText)
xLocker.BoundText = IIf(xRelation.BoundText <> "" Or xYacht.BoundText <> "" Or xMeeting.Value <> 0, "", xLocker.BoundText)
xYacht.BoundText = IIf(xRelation.BoundText <> "" Or xLocker.BoundText <> "" Or xMeeting.Value <> 0, "", xYacht.BoundText)
'xMeeting.Value = IIf(xRelation.BoundText <> "" Or xLocker.BoundText <> "" Or xYacht.BoundText <> "", 0, xMeeting.Value)
xDays.Text = IIf(xRelation.BoundText <> "" Or xLocker.BoundText <> "" Or xYacht.BoundText <> "", "", xDays.Text)

CardTable.FindFirst "Code = " & xCode.Text
If CardTable.NoMatch Then
    CardTable.AddNew
Else
    CardTable.Edit
End If
CardTable!CODE = xCode.Text
CardTable!DESCA = TurnValue(xDescA, "", Null)
CardTable!Value = Val(Format(xValue.Text, "##,##0.00"))
CardTable!Relation = TurnValue(xRelation.BoundText, "", Null)
CardTable!isMember = xIsMember.Value
CardTable!age1 = TurnValue(Val(xAge1.Text), 0, Null)
CardTable!age2 = TurnValue(Val(xAge2.Text), 0, Null)
CardTable!Sex = ReplaceSex

CardTable!Locker = TurnValue(xLocker.BoundText, "", Null)
CardTable!yacht = TurnValue(xYacht.BoundText, "", Null)
CardTable!Meeting = xMeeting.Value
CardTable!BASICOLD = xBasicOld.Value
CardTable!showpaid = xShowPaid.Value
CardTable!days = TurnValue(Val(xDays.Text), 0, Null)

CardTable!NoRate = xNoRate.Value
CardTable!BasicNew = xBasicNew.Value
CardTable!BasicDied = xBasicDied.Value
CardTable!LATE = xLate.Value
CardTable!AllMember = xAllMember.Value
CardTable.Update
If xAllYears.Value = 0 Then myreplaceGrd Else MyReplaceGrdall
Me.MousePointer = 0
'MsgBox " „ «·Õ›Ÿ »‰Ã«Õ"

End Sub
Function myvalid() As Boolean
If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·»‰œ €Ì— „”Ã·"
    Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "»Ì«‰ «·»‰œ €Ì— „”Ã·"
    Exit Function
End If
myvalid = True
End Function
Private Sub CmdAdd_Click()
CardTable.MoveLast
xCode.Text = CardTable.CODE + 1
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
If Not myValidDelete Then Exit Sub
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
mydb.Execute "Delete * From FILE2_11 Where item = " & xCode.Text
mydb.Execute "Delete * From FILE2_10 Where Code = " & xCode.Text
CardTable.Requery
If CardTable.RecordCount > 0 Then
    CardTable.FindLast "Code < " & xCode.Text
    If CardTable.NoMatch Then CardTable.MoveFirst
    myload
Else
    myDefine
End If
End If
End Sub
Private Sub CmdExit_Click()
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
Private Sub cmdSave_Click()
'msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not myvalid Then Exit Sub

'If Not MsgBox(msgBoxStr, 4) = 6 Then
'    CmdUndo_Click
'    Exit Sub
'End If
myreplace
Inform " „  «·Õ›Ÿ »‰Ã«Õ"
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.FindFirst "code = " & xCode.Text
    myload
End If

End Sub
Private Sub cmdundo_Click()
If CardTable.RecordCount = 0 Then
    myDefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
        myload
    Else
        myload
    End If
End If
End Sub
Private Sub MyReplaceGrdall()
With grid1
For nYear = xYear.ListIndex To xYear.ListCount - 1
    mydb.Execute "DELETE * FROM FILE2_40 WHERE ITEM = " & xCode.Text & " AND [YEAR] = " & xYear.ItemData(nYear)
    For nRow = 2 To grid1.rows - 1
        For nCol = 4 To grid1.Cols - 1
            mydb.Execute "INSERT INTO FILE2_40(ITEM,[YEAR],[TYPE],SECTION,[VALUE],DISCOUNT,BASIC) " & _
                        " VALUES ( " & _
                        addvalue(xCode.Text) & "," & _
                        addvalue(xYear.ItemData(nYear)) & "," & _
                        addvalue(.TextMatrix(0, nCol)) & "," & _
                        addstring(.TextMatrix(nRow, 0)) & "," & _
                        addvalue(.TextMatrix(nRow, 2)) & "," & _
                        addvalue(.TextMatrix(nRow, 3)) & "," & _
                        IIf(Val(.TextMatrix(nRow, nCol)) = 0, "FALSE", "TRUE") & ")"
        Next
    Next
Next
End With
End Sub
Private Sub Command1_Click()
Dim printTable As Recordset
Set printTable = mydb.OpenRecordset("SELECT * FROM FILE2_10 ORDER BY CODE", dbOpenDynaset)
mydb.Execute "delete * from temp"
Do Until printTable.EOF
    mydb.Execute "Insert Into Temp(Str1,Str2,Str3) " & _
                 "values(" & _
                 addstring(printTable!CODE) & "," & _
                 addstring(Chr(254) & Replace(printTable!DESCA & "", " ", Chr(254) & " " & Chr(254))) & "," & _
                 addstring(printTable!Value) & _
                 ")"
    printTable.MoveNext
Loop
myws.BeginTrans
myws.CommitTrans
Report1.Reset
Report1.WindowShowExportBtn = False
Report1.ReportFileName = App.Path & "\RPT\items.rpt"
Report1.DataFiles(0) = App.Path & "\MDB\DATA.mdb"
Report1.Action = 1: tempdb.Execute "Delete * from temp"
printTable.Close
Set printtabe = Nothing
End Sub

Private Sub Command2_Click()
For i = 3 To grid1.Cols - 1
    grid1.ColHidden(i) = True
Next
PrintGrdNew.doprint grid1, 1.5, 0, "ÿ»«⁄… »‰œ : " & xDescA.Text, "„Ê”„ " & xYear.Text, , , False, False, 10
PrintGrdNew.Show 1
For i = 3 To grid1.Cols - 1
    grid1.ColHidden(i) = False
Next
End Sub

Private Sub Form_Load()
Set CardTable = mydb.OpenRecordset("SELECT * FROM FILE2_10 ORDER BY CODE", dbOpenDynaset)
Set SectionItemTable = mydb.OpenRecordset("FILE2_11", dbOpenDynaset)
Set Deletetable1 = mydb.OpenRecordset("FILE2_30", dbOpenSnapshot)
Set RecordCountTable = mydb.OpenRecordset("FILE2_10")
For i = 0 To Val(Format(Date, "yy"))
   xYear.AddItem paidYearString(2000 + i)
   xYear.ItemData(i) = 2000 + i
Next
xYear.Text = paidYearString(PaidYear(Date))

Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA from FILE0_00 WHERE FLAG = 0 order by Code"
xRelation.BoundColumn = "Code"
xRelation.ListField = "DESCA"

data2.DatabaseName = MdbPath
data2.RecordSource = "Select Code,DescA from file0_00 Where Flag = 6"
xLocker.BoundColumn = "Code"
xLocker.ListField = "DESCA"

Data3.DatabaseName = MdbPath
Data3.RecordSource = "Select Code,DescA from file0_00 Where Flag = 8"
xYacht.BoundColumn = "Code"
xYacht.ListField = "DESCA"
FixGrid

If CardTable.RecordCount > 0 Then
    CardTable.MoveLast
    myload
Else
    xCode.Text = 1
    myDefine
End If
End Sub

Private Sub lblYear_Change()
LoadGrd
End Sub
Private Sub xCode_LostFocus()
If xCode.Text = "" Then Exit Sub
CardTable.FindFirst "Code = " & xCode.Text
If Not CardTable.NoMatch Then myload
End Sub
Private Function myRecordCount() As Integer
If RecordCountTable.RecordCount = 0 Then Exit Function
RecordCountTable.MoveLast
myRecordCount = RecordCountTable.RecordCount
End Function
Private Function myValidDelete() As Boolean
Deletetable1.FindFirst "Item = " & xCode.Text
If Not Deletetable1.NoMatch Then
    MsgBox "«·»‰œ „”Ã· ›Ï „ÿ«·»«  ”«»ﬁ… ÌÃ» Õ–› „‰ Â–Â «·„ÿ«·»« "
    Exit Function
End If
myValidDelete = True
End Function
Private Sub xLocker_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xLocker.BoundText = ""
End Sub
Private Sub xRelation_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xRelation.BoundText = ""
End Sub
Private Sub xYacht_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xYacht.BoundText = ""
End Sub
Private Function LoadSex(cSex) As Integer
If IsNull(cSex) Then Exit Function
If cSex = 0 Then LoadSex = 1
If cSex = 1 Then LoadSex = 2
End Function
Private Function ReplaceSex()
If xSex.ListIndex = 0 Then
    RelpaceSex = Null
    Exit Function
End If
ReplaceSex = xSex.ListIndex - 1
End Function
Private Sub FixGrid()
Dim paidTypeTable As Recordset, SectionTable As Recordset
Set SectionTable = mydb.OpenRecordset("Select * From File0_10 order by Code", dbOpenSnapshot)
Set paidTypeTable = mydb.OpenRecordset("Select * From File0_50  where code >= 0 order by Code")
paidTypeTable.MoveLast
With grid1
.rows = 2
.ColHidden(0) = True
.RowHidden(0) = True
paidTypeTable.MoveLast
.Cols = paidTypeTable.RecordCount + 4
.TextMatrix(1, 1) = "«·›∆…"
.TextMatrix(1, 2) = "«·ﬁÌ„…"
.TextMatrix(1, 3) = "«·Œ’„"

i = 4
paidTypeTable.MoveFirst
Do
    grid1.TextMatrix(0, i) = paidTypeTable.CODE
    grid1.TextMatrix(1, i) = TurnValue(paidTypeTable.desca2, "", Null)
    .ColWidth(i) = 900
    .ColDataType(i) = flexDTBoolean
    paidTypeTable.MoveNext
    i = i + 1
Loop Until paidTypeTable.EOF

.ColWidth(1) = 2000
.ColWidth(2) = 750
.ColWidth(3) = 750
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
    nWidth = nWidth + grid1.ColWidth(i)
Next
If grid1.Width > nWidth Then
    grid1.Left = grid1.Left + (grid1.Width - nWidth)
    grid1.Width = nWidth
End If


SectionTable.MoveFirst
Do
    .AddItem ""
    .TextMatrix(.rows - 1, 0) = SectionTable!CODE
    .TextMatrix(.rows - 1, 1) = SectionTable!DESCA
    SectionTable.MoveNext
Loop Until SectionTable.EOF
End With
End Sub
Private Sub LoadGrd()
Set GRDTABLE = mydb.OpenRecordset("Select * From File2_40 " & _
               " where Item = " & xCode.Text & _
               " and [year] = " & xYear.ItemData(xYear.ListIndex))

With grid1
For nRow = 2 To grid1.rows - 1
    .TextMatrix(nRow, 2) = ""
    .TextMatrix(nRow, 3) = ""
     For nCol = 4 To grid1.Cols - 1
         .TextMatrix(nRow, nCol) = ""
     Next
Next

If GRDTABLE.RecordCount = 0 Then
    Exit Sub
End If

For nRow = 2 To grid1.rows - 1
    For nCol = 4 To grid1.Cols - 1
        GRDTABLE.FindFirst "Section = " & MyParn(grid1.TextMatrix(nRow, 0)) & _
                           " AND [TYPE] = " & grid1.TextMatrix(0, nCol)
        If Not GRDTABLE.NoMatch Then
            .TextMatrix(nRow, 2) = TurnValue(GRDTABLE!Value, Null, "")
            .TextMatrix(nRow, 3) = TurnValue(GRDTABLE!Discount, Null, "")
            .TextMatrix(nRow, nCol) = IIf(GRDTABLE!basic, -1, 0)
        End If
    Next
Next
End With
End Sub
Private Sub myreplaceGrd()
With grid1
mydb.Execute "DELETE * FROM FILE2_40 WHERE YEAR = " & xYear.ItemData(xYear.ListIndex) & _
             " AND ITEM = " & xCode.Text

For nRow = 2 To grid1.rows - 1
    For nCol = 4 To grid1.Cols - 1
        mydb.Execute "INSERT INTO FILE2_40(ITEM,[YEAR],[TYPE],SECTION,[VALUE],DISCOUNT,BASIC) " & _
                    " VALUES ( " & _
                    addvalue(xCode.Text) & "," & _
                    addvalue(xYear.ItemData(xYear.ListIndex)) & "," & _
                    addvalue(.TextMatrix(0, nCol)) & "," & _
                    addstring(.TextMatrix(nRow, 0)) & "," & _
                    addvalue(.TextMatrix(nRow, 2)) & "," & _
                    addvalue(.TextMatrix(nRow, 3)) & "," & _
                    IIf(Val(.TextMatrix(nRow, nCol)) = 0, "FALSE", "TRUE") & ")"
    Next
Next
End With
End Sub
Private Sub xYear_Click()
If xCode.Text = "" Then Exit Sub
LblYear.Caption = xYear.ItemData(xYear.ListIndex)
End Sub
