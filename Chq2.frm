VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form chq2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "‘Ìﬂ« "
   ClientHeight    =   6615
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9210
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
   ScaleHeight     =   6615
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox XF_PAY 
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
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3169
      Width           =   1365
   End
   Begin VB.TextBox xBal 
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
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3536
      Width           =   1365
   End
   Begin VB.CommandButton cmd_subsave 
      BackColor       =   &H00CAD29F&
      Caption         =   "Õ›‹‹‹‹‹‹‹‹‹‹‹Ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2715
   End
   Begin VB.Data Data1 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   675
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.OptionButton xClosed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "—›÷ / —œ «·‘Ìﬂ"
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
      Height          =   345
      Index           =   1
      Left            =   3750
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4875
      Width           =   1740
   End
   Begin VB.OptionButton xClosed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "„Õ’·"
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
      Height          =   345
      Index           =   2
      Left            =   5850
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4875
      Width           =   1065
   End
   Begin VB.OptionButton xClosed 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "‘Ìﬂ €Ì— „Õ’·"
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
      Height          =   345
      Index           =   0
      Left            =   7200
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   4875
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9210
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   9210
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4065
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0FFFF&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3150
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   900
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1650
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9210
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6075
      Width           =   9210
      Begin VB.OptionButton optclose 
         BackColor       =   &H00800000&
         Caption         =   "„—›Ê÷…"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   4950
         TabIndex        =   23
         Top             =   150
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H00800000&
         Caption         =   "„Õ’·…"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   6450
         TabIndex        =   22
         Top             =   150
         Width           =   990
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H00800000&
         Caption         =   "€Ì— „Õ’·…"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   21
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton CMDNEXT 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   75
         Width           =   1065
      End
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
      Left            =   5850
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1334
      Width           =   1215
   End
   Begin VB.TextBox xNAME1 
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
      Left            =   1800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   2715
   End
   Begin VB.TextBox xDATE_1 
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
      Left            =   4350
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2435
      Width           =   2715
   End
   Begin VB.TextBox xBANK_REC 
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
      Left            =   4350
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2068
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox xNAME2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4350
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1701
      Width           =   2715
   End
   Begin VB.TextBox xMEMO 
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
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5700
      Width           =   4965
   End
   Begin VB.TextBox xDATE_3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4725
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5325
      Width           =   2340
   End
   Begin VB.TextBox xValue 
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
      Left            =   5700
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2802
      Width           =   1365
   End
   Begin VB.TextBox XSER_NO 
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
      Left            =   4725
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2340
   End
   Begin VB.TextBox XCHK_ID 
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
      Left            =   4725
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   967
      Width           =   2340
   End
   Begin VB.TextBox xDATE_R 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4350
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3903
      Width           =   2715
   End
   Begin MSDBCtls.DBCombo XID_BANK 
      DataSource      =   "Data2"
      Height          =   315
      Left            =   4350
      TabIndex        =   9
      Top             =   4275
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid VsPay 
      Height          =   3015
      Left            =   150
      TabIndex        =   35
      Top             =   1725
      Width           =   3090
      _cx             =   5450
      _cy             =   5318
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   13292191
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   16777215
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
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Chq2.frx":0000
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   4
   End
   Begin MSDBCtls.DBCombo xBox 
      Bindings        =   "Chq2.frx":0088
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2100
      TabIndex        =   49
      Top             =   5325
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Œ“‰…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4215
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ  Õ—Ì— :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   3915
      Width           =   975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”· ‘Ìﬂ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   675
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·‘Ìﬂ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label LabelCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·„Ê—œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1395
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   2475
      Width           =   1290
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   2115
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„ŸÂ— „‹‹‹‹‹‰"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1755
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»‰ﬂ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   4275
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "€Ì— „”œœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3555
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "”œ«œ ”«»ﬁ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3195
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﬂÊœ «·„Ê—œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   75
      Width           =   795
   End
   Begin VB.Label LabelName1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "≈”„ «·„Ê—œ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4725
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1350
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «· Õ’Ì·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5400
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ«  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5775
      Width           =   735
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„… :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2835
      Width           =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   4
      X1              =   3450
      X2              =   8925
      Y1              =   4725
      Y2              =   4725
   End
End
Attribute VB_Name = "chq2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UPDATE FILE5_21 SET FILE5_21.CLOSED = "2" WHERE (((FILE5_21.DATE_3) Is Not Null));
'
Dim cardtable As Recordset, RecordCountTable As Recordset
Dim cFileName As String
Dim alltable As Recordset, ClientTable1 As Recordset
Dim cClient1 As String, cClient2 As String
Dim cFileMove1 As String, cFileMove2 As String, cChqDesc As String
Dim cFieldUnder As String, cFieldTrans, cFieldReject As String
Dim CMOVE As String
Dim PayChqTable As Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value)
CmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value)
CmdUndo.Enabled = (nMode = LoadMode Or optclose(0).Value)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
CMDNEXT.Enabled = (nMode = LoadMode)
CmdLast.Enabled = (nMode = LoadMode)
CmdFirst.Enabled = (nMode = LoadMode)
XSER_NO.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(3)
Dim GrdArray(5)
Set Generalarray(1) = Me
If optclose(0).Value Then cWhere = " WHERE CLOSED = '0' "
If optclose(1).Value Then cWhere = " WHERE CLOSED = '1'"
If optclose(2).Value Then cWhere = " WHERE CLOSED = '2'"
cStr1 = " SELECT FILE5_21.SER_NO AS „”·”· , FILE5_21.NAME1 AS ⁄„Ì· , FILE5_21.NAME2 AS „ŸÂ— , format(FILE5_21.DATE_1 ,'d-m-yyyy') AS ≈” Õﬁ«ﬁ, format(FILE5_21.DATE_r ,'d-m-yyyy') AS  Õ—Ì—"
Generalarray(2) = " SELECT FILE5_21.SER_NO AS „”·”· , FILE5_21.NAME1 AS ⁄„Ì· , FILE5_21.NAME2 AS „ŸÂ— ,format(FILE5_21.DATE_1 ,'d-m-yyyy') AS ≈” Õﬁ«ﬁ, format(FILE5_21.DATE_r ,'d-m-yyyy') AS  Õ—Ì— From " & cFileName & cWhere
Generalarray(3) = " AND ( NAME1 Like '*cFilter*' OR NAME2 Like '*cFilter*' ) "
       
GrdArray(1) = 1300
GrdArray(2) = 2000
GrdArray(3) = 2000
GrdArray(4) = 1500
GrdArray(5) = 1500
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Sub ClientLookup(pFlag)
Dim Generalarray(3)
Dim GrdArray(2)
Set Generalarray(1) = Me
If pFlag = 1 Then
    Generalarray(2) = "Select Code as [«·ﬂÊœ] ,Desca as [«·≈”„] From " & cClient1
Else
    Generalarray(2) = "Select Code as [«·ﬂÊœ] ,Desca as [«·≈”„] From " & cClient2
End If
Generalarray(3) = " Where DESCA Like '*cFilter*'"
       
GrdArray(1) = 1200
GrdArray(2) = 4000
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Sub myDefine()
XCHK_ID.Text = ""
XSER_NO.Text = ""

xCode.Text = ""
xNAME1.Text = ""

xNAME2.Text = ""
XID_BANK.BoundText = ""
xNAME1.Text = ""
xBox.BoundText = ""
xNAME2.Text = ""
xBANK_REC.Text = ""
xDATE_1.Text = ""
xDATE_3.Text = ""
xDATE_R.Text = ""
xValue.Text = ""
XF_PAY.Text = ""
xMEMO.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False

Handlecontrols DefineMode
If myRecordCount = 0 Then
    xRecordNumber = ""
Else
    xRecordNumber = "”Ã· " & myRecordCount + 1 & " „‰ " & myRecordCount + 1
End If
End Sub
Sub myProc()
cardtable.FindFirst "SER_NO = " & MyParn(GrdText(Search.grid1, 0))
If TypeOf ActiveControl Is TextBox Then
'    If ActiveControl.Name = Me.xCode.Name Then xCode.Text = GrdText(Search.Grid1, 0)
'    If ActiveControl.Name = Me.xCode2.Name Then xCode2.Text = GrdText(Search.Grid1, 0)
    ActiveControl.Text = GrdText(Search.grid1, 0)
'    ActiveControl.Text = GrdText(Search.Grid1, 0)
Else
    myload
End If
Unload Search
End Sub
Sub myload()
Dim nBal As Double
nBal = 0
XSER_NO.Text = TurnValue(cardtable.ser_no, Null, "")
XCHK_ID.Text = TurnValue(cardtable.chk_id, Null, "")
xCode.Text = TurnValue(cardtable.CODE, Null, "")

'xSer_no2.Text = TurnValue(CardTable.code2, Null, "")


xClosed(0).Value = IIf(cardtable.CLOSED = "0", True, False)
xClosed(1).Value = IIf(cardtable.CLOSED = "1", True, False)
xClosed(2).Value = IIf(cardtable.CLOSED = "2", True, False)



If xCode.Text <> "" Then
    xCode_LostFocus
Else
    xNAME1.Text = ""
End If
xNAME2.Text = TurnValue(cardtable.name2, Null, "")
xBANK_REC.Text = TurnValue(cardtable.BANK_REC, Null, "")
xValue.Text = Format(TurnValue(cardtable.Value, Null, 0), "#0.00")
XF_PAY.Text = Format(TurnValue(cardtable.F_PAY, Null, 0), "#0.00")
xDATE_1.Text = Format(cardtable.date_1, "dd-mm-yyyy")
xDATE_3.Text = Format(cardtable.date_3, "dd-mm-yyyy")
xDATE_R.Text = Format(cardtable.date_r, "dd-mm-yyyy")
xMEMO.Text = TurnValue(cardtable.Memo, Null, "")
xBox.BoundText = TurnValue(cardtable.BOX, Null, "")
XID_BANK.BoundText = TurnValue(cardtable.ID_BANK, Null, "")

'xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & myRecordCount
With VsPay
    .Rows = 1
    PayChqTable.FindFirst " ser_no = " & MyParn(XSER_NO.Text)
    If Not PayChqTable.NoMatch Then
        Do While Not PayChqTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = Format(PayChqTable.Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 1) = Format(PayChqTable.Value, "#0.00")
            nBal = nBal + Val(.TextMatrix(.Rows - 1, 1))
            PayChqTable.MoveNext
            If PayChqTable.EOF Then Exit Do
            If PayChqTable.ser_no <> XSER_NO.Text Then Exit Do
        Loop
    Else
        .Rows = 2
    End If
End With
xBal.Text = Format((xValue.Text) - nBal, "#0.00")
xBal.Visible = Not (Val(xBal.Text) = 0)
Handlecontrols LoadMode
End Sub
Function MYVALID() As Boolean
If XSER_NO.Text = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· „””·”· ··‘Ìﬂ"
    Exit Function
End If
If Not IsDate(xDATE_R.Text) Then
    MsgBox "ÌÃ»  ”ÃÌ·  «—ÌŒ «· Õ—Ì—"
    Exit Function
End If
If xNAME1.Text = "" Then
    MsgBox "«·≈”„  ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If
If xCode.Text <> "" Then
    ClientTable1.FindFirst " CODE = " & MyParn(xCode.Text)
    If ClientTable1.NoMatch Then Exit Function
End If
If Not xClosed(0).Value And Not IsDate(xDATE_3.Text) Then
    MsgBox " ”ÃÌ·  «—ÌŒ «· ŸÂÌ— °  ÕœÌœ «‰ «·‘Ìﬂ  „  ŸÂÌ—…"
    Exit Function
End If
MYVALID = True
End Function

Private Sub cmd_subsave_Click()
Dim nTot As Double
nTot = 0

With VsPay
    mydb.Execute "delete * From file5_23 where Ser_No = " & MyParn(XSER_NO.Text)
    For i = 1 To .Rows - 1
        If IsDate(.TextMatrix(i, 0)) And Val(.TextMatrix(i, 1)) <> 0 Then
            PayChqTable.AddNew
            PayChqTable.ser_no = XSER_NO.Text
            PayChqTable.Date = .TextMatrix(i, 0)
            PayChqTable.Value = Val(.TextMatrix(i, 1))
            nTot = nTot + Val(.TextMatrix(i, 1))
            PayChqTable.Update
        End If
    Next i
    If nTot = Val(xValue.Text) + Val(XF_PAY.Text) Then
        If IsDate(VsPay.TextMatrix(.Rows - 1, 0)) Then
            xDATE_3.Text = VsPay.TextMatrix(.Rows - 1, 0)
        Else
            xDATE_3.Text = VsPay.TextMatrix(.Rows - 2, 0)
        End If
        xClosed(2).Value = True
    End If
End With
cmdSave_Click
End Sub
Private Sub CmdAdd_Click()
myDefine
If alltable.RecordCount = 0 Then
    XSER_NO.Text = "000001"
Else
    alltable.MoveLast
    XSER_NO.Text = IncRec(myLastField(alltable, "Ser_No"))
End If
'XSER_NO.SetFocus
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    mydb.Execute "delete * From " & cFileName & " where Ser_No = " & MyParn(XSER_NO.Text)
    mydb.Execute "Delete * from " & cFileMove1 & " where Doc_ID = " & MyParn(XSER_NO.Text) & " AND ([TYPE] = 'A' OR [TYPE] = 'C' OR [TYPE] = 'F')"
    mydb.Execute "Delete * from " & cFileMove2 & " where Doc_ID = " & MyParn(XSER_NO.Text) & " AND ([TYPE] = 'B' OR [TYPE] = 'D' OR [TYPE] = 'E')"
End If
alltable.Requery
cardtable.Requery
If cardtable.RecordCount > 0 Then
    cardtable.FindLast "SER_NO < " & MyParn(XSER_NO.Text)
    If cardtable.NoMatch Then cardtable.MoveFirst
    myload
Else
    If optclose(0).Value Then CmdAdd_Click Else myDefine
End If
End Sub
Private Sub cmdexit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
cardtable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
cardtable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
cardtable.MoveNext
If cardtable.EOF Then
    cardtable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
cardtable.MovePrevious
If cardtable.BOF Then
    cardtable.MoveNext
Else
    myload
End If
End Sub
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub

If Not MsgBox(msgBoxStr, 4) = 6 Then
    CmdUndo_Click
    Exit Sub
End If
MyReplace
cardtable.Requery

If XSER_NO.Enabled Then
    CmdAdd_Click
Else
    If cardtable.RecordCount > 0 Then
        cardtable.FindFirst "SER_NO = " & MyParn(XSER_NO.Text)
        If cardtable.NoMatch Then
            cardtable.FindLast "SER_NO <= " & MyParn(XSER_NO.Text)
            If cardtable.NoMatch Then cardtable.MoveFirst
        End If
        myload
    Else
        myDefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
If cardtable.RecordCount = 0 Then
    myDefine
Else
    If XSER_NO.Enabled Then
        cardtable.MoveLast
        myload
    Else
        myload
    End If
End If
End Sub
Private Sub Form_Load()
    Me.Caption = "√Ê—«ﬁ ﬁ»÷"
    cChqDesc = "√Ê—«ﬁ ﬁ»÷"
    cFileName = "File5_21"
    cClient1 = "File4_10"
    cClient2 = "File4_10"
    cFileMove1 = "File4_11"
    cFileMove2 = "File4_11"
    Set PayChqTable = mydb.OpenRecordset("select * from file5_23 order by ser_no , date ")


Set alltable = mydb.OpenRecordset("Select * from " & cFileName & "  order by Ser_No")
Set ClientTable1 = mydb.OpenRecordset(cClient1, dbOpenDynaset)

Data2.DatabaseName = MdbPath
Data2.RecordSource = "SELECT CODE, DESCA FROM FILE1_70 WHERE FLAG = 6 "
XID_BANK.BoundColumn = "CODE"
XID_BANK.ListField = "DESCA"

Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE0_50"
xBox.BoundColumn = "CODE"
xBox.ListField = "DESCA"


optclose(0).Value = True
With VsPay
    .Cols = 2
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = " «—ÌŒ"
    .TextMatrix(0, 1) = "«·ﬁÌ„…"
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .Editable = flexEDKbdMouse
End With
If cardtable.RecordCount > 0 Then
    cardtable.MoveLast
    myload
Else
    CmdAdd_Click
End If
End Sub
Private Sub VsPay_EnterCell()
    If VsPay.row = VsPay.Rows - 1 And IsDate(VsPay.TextMatrix(VsPay.row, 0)) Then VsPay.AddItem ""
End Sub
Private Sub xClosed_Click(Index As Integer)
If Index = 0 Then
    xDATE_3.Text = ""
End If
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookup 1

End Sub

Private Sub xCode_LostFocus()
ClientTable1.FindFirst "code = " & MyParn(xCode.Text)
If Not ClientTable1.NoMatch Then
    xNAME1.Text = ClientTable1!DESCA
End If
End Sub

Private Sub xCode2_Change()
If xCode2.Text <> "" Then
    xCode.Text = ""
    xCode.Enabled = False
    xNAME1.Text = ""
Else
    xCode.Enabled = True
End If
End Sub
Private Sub xSer_no_LostFocus()
cardtable.FindFirst "SER_NO = " & MyParn(XSER_NO.Text)
If Not cardtable.NoMatch Then myload
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
Sub MyReplace()
cardtable.FindFirst "SER_NO = " & MyParn(XSER_NO.Text)
If cardtable.NoMatch Then
    cardtable.AddNew
Else
    cardtable.Edit
End If
cardtable.ser_no = TurnValue(XSER_NO.Text, "", Null)
cardtable.chk_id = TurnValue(XCHK_ID.Text, "", Null)
cardtable.CODE = TurnValue(xCode.Text, "", Null)
cardtable.name1 = TurnValue(xNAME1.Text, "", Null)
cardtable.name2 = TurnValue(xNAME2.Text, "", Null)
cardtable.BANK_REC = TurnValue(xBANK_REC.Text, "", Null)
'cardtable.date_1 = DateFix(xDATE_1.Text)
cardtable.date_3 = DateFix(xDATE_3.Text)
cardtable.date_r = DateFix(xDATE_R.Text)
cardtable.Value = Val(xValue.Text)
cardtable.F_PAY = Val(XF_PAY.Text)
cardtable.Memo = TurnValue(xMEMO.Text, "", Null)
cardtable.ID_BANK = TurnValue(XID_BANK.BoundText, "", Null)
cardtable.BOX = TurnValue(xBox.BoundText, "", Null)

If xClosed(0).Value Then
    cardtable.CLOSED = 0
End If
If xClosed(1).Value Then
    cardtable.CLOSED = 1
End If
If xClosed(2).Value Then
    cardtable.CLOSED = 2
End If
cardtable.Update
cardtable.Requery
alltable.Requery
MYREPLACEMOVE
End Sub
Private Sub Cmd_Trans_Click()
If IsDate(xDATE_3.Text) Then
    cmdSave_Click
    If XSER_NO.Text <> "" And xTransCode1.Text <> "" Then
        mydb.Execute " UPDATE FILE5_21 SET FILE5_21.CLOSED  = '2' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        mydb.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
        cString = "Insert Into File4_11(" & _
                  "[Type],Doc_Id,Code,[Date],PAY,DescA)" & _
                  " Select 'T',Ser_No, TRANSCODE1 ,Date_3,Value,'  ŸÂÌ— ‘Ìﬂ Õﬁ ' & date_1" & _
                  " From File5_21" & _
                  " WHERE FILE5_21.SER_NO  = " & MyParn(XSER_NO.Text)
        mydb.Execute cString
    End If
    
    If xCode2.Text <> "" And xTransCode2.Text <> "" Then
        mydb.Execute " UPDATE FILE5_21 SET FILE5_21.CLOSED  = '2' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        cmdSave_Click
        mydb.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
        cString = "Insert Into File4_11(" & _
                  "[Type],Doc_Id,Code,[Date],SAL,DescA)" & _
                  " Select 'T',Ser_No,TRANSCODE2 ,Date_3,Value,'  ŸÂÌ— ‘Ìﬂ Õﬁ ' & date_1" & _
                  " From File5_21" & _
                  " WHERE FILE5_21.SER_NO  = " & MyParn(XSER_NO.Text)
        mydb.Execute cString
    End If
Else
    MsgBox " ”ÃÌ·  «—ÌŒ  ŸÂÌ— «·‘Ìﬂ "
End If
End Sub
Private Sub Cmd_unTrans_Click()
If IsDate(xDATE_3.Text) Then
    xTransCode1.Text = ""
    xTransCode2.Text = ""
    cmdSave_Click
    If xCode.Text <> "" And xTransCode1.Text <> "" Then
        mydb.Execute " UPDATE FILE5_21 SET FILE5_21.CLOSED  = '0' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        mydb.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
    End If
    
    If xCode2.Text <> "" And xTransCode2.Text <> "" Then
        mydb.Execute " UPDATE FILE5_21 SET FILE5_21.CLOSED  = '0' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        mydb.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
    End If
End If
End Sub
Private Sub CmdColect_Click()
xBox.Enabled = True
xDATE_3.Enabled = True
xMEMO.Enabled = True
xDATE_3.SetFocus
End Sub
Private Sub CmdDelColect_Click()
If MsgBox("Â·  Êœ«·€«¡  Õ’Ì· «·‘Ìﬂ", vbOKCancel, "«·„—‘œ") = vbCancel Then Exit Sub
xDATE_3.Text = ""
xMEMO.Text = ""
cardtable.Edit
cardtable.date_3 = Null
cardtable.CLOSED = "0"
cardtable.Memo = Null
cardtable.Update
cardtable.Requery

If cardtable.RecordCount = 0 Then
    optclose(2).Enabled = False
    optclose(0).Value = True
Else
    cardtable.MoveLast
    myload
End If
End Sub
Private Sub MYREPLACEMOVE()
    mydb.Execute "Delete * from file4_11 where Doc_ID = " & MyParn(XSER_NO.Text) & " AND [TYPE] = '8' "
    mydb.Execute "Delete * from file4_11 where Doc_ID = " & MyParn(XSER_NO.Text) & " AND [TYPE] = '9' "
    
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],Pay,DescA,SHOW )" & _
              " Select '8',SER_NO,Code,[Date_R],[Value],'‘Ìﬂ Õﬁ' & FORMAT(DATE_1,'dd-mm-yyyy') , '2' " & _
              "  From File5_21 WHERE SER_NO = " & MyParn(XSER_NO.Text)
    mydb.Execute cString

    cString = "INSERT INTO FILE4_11 ( PAY, TYPE, SHOW, DOC_ID, CODE, [DATE], DESCA )" & _
              "SELECT file5_23.VALUE, '9'  AS Expr1, 3 AS Expr2, file5_23.SER_NO, FILE5_21.CODE, FILE5_23.DATE, 'œ›⁄… ‘Ìﬂ' " & _
              " FROM FILE5_21 RIGHT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no " & _
              " WHERE FILE5_23.SER_NO = " & MyParn(XSER_NO.Text)
    mydb.Execute cString

End Sub
Private Sub Command1_Click()
    mydb.Execute "Delete * From File4_11 WHERE [TYPE] = '9' "
        
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],PAY,DescA)" & _
              " Select '9',Ser_No,Code,Date_R,Value,' √Ê—«ﬁ ﬁ»÷ Õﬁ ' & date_1" & _
              " From File5_21" & _
              " WHERE Not FILE5_21!old AND FILE5_21.CODE IS NOT NULL "
    mydb.Execute cString
        
    mydb.Execute "Delete * From File4_11 WHERE [TYPE] = '8' "
    cString = "Insert Into File4_11(" & _
              "[Type],Doc_Id,Code,[Date],SAL,DescA)" & _
              " Select '8',Ser_No,Code2,Date_R,Value,' √Ê—«ﬁ ﬁ»÷ Õﬁ ' & date_1" & _
              " From File5_21" & _
              " WHERE Not FILE5_21!old  AND FILE5_21.CODE2 IS NOT NULL"
    mydb.Execute cString

MsgBox " „  ⁄„·Ì… «·÷»ÿ »‰Ã«Õ", , "«·‘Ìﬂ« "
End Sub

Private Sub OptClose_Click(Index As Integer)
cString = "Select * From " & cFileName & " Where closed = " & MyParn(Index) & " Order by Ser_No "
Set cardtable = mydb.OpenRecordset(cString, dbOpenDynaset)
If cardtable.RecordCount > 0 Then
    cardtable.MoveLast
    myload
Else
    If optclose(0).Value Then
        CmdAdd_Click
    Else
        myDefine
    End If
End If
End Sub
Private Sub xSer_no_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê  From " & cFile_10
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 2600
    GrdArray(3) = 1500
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub handleopt()
alltable.Requery
alltable.FindFirst "closed = '1'"
optclose(1).Enabled = Not alltable.NoMatch
If Not optclose(1).Enabled And optclose(1).Value = True Then optclose(0).Value = True

alltable.FindFirst "closed = '2'"
optclose(2).Enabled = Not alltable.NoMatch
If Not optclose(2).Enabled And optclose(2).Value = True Then optclose(0).Value = True
End Sub
Private Sub ChqLookup()
Dim Generalarray(4)
Dim GrdArray(5)
Set Generalarray(1) = Me
Generalarray(2) = "Select Ser_No as [„”·”· «·‘Ìﬂ],Name1 As [„Ê—œ ],Name12 As [⁄„Ì·],FORMAT(Date_1,'DD-MM-YYYY') As [ «—ÌŒ «” Õﬁ«ﬁ «·‘Ìﬂ],Value as [ﬁÌ„… «·‘Ìﬂ] From " & cFileName & "  Where Closed = " & MyParn(sClose)
Generalarray(3) = " and ( Name1 Like '*cFilter*' Or NAME12 Like '*cFilter*') "
Generalarray(4) = "Order By Name1 , name12 "
   
GrdArray(1) = 1000
GrdArray(2) = 2000
GrdArray(3) = 2000
GrdArray(4) = 1200
GrdArray(5) = 1200

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„ "
Search.Show 1
End Sub
Private Sub SetCardTable()
cString = "Select * From " & cFileName & " Where closed = " & MyParn(xType.ListIndex) & " Order by Ser_No "
If cardtable.RecordCount > 0 Then
    cardtable.MoveLast
    myload
Else
    XSER_NO.Text = ""
End If
End Sub
