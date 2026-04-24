VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPaidItem2 
   Caption         =   "»‰Êœ «·«‘ —«ﬂ« "
   ClientHeight    =   7935
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
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "ÿ»«⁄… »‰Êœ «·⁄«„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   4005
      Width           =   1815
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   2865
      Left            =   75
      TabIndex        =   42
      Top             =   4425
      Width           =   11715
      _cx             =   20664
      _cy             =   5054
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
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   2205
      Width           =   3030
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
         TabIndex        =   47
         Top             =   315
         Width           =   1830
      End
   End
   Begin VB.CheckBox xAllYears 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   " ÿ»Ìﬁ ⁄·Ì ﬂ· «·”‰Ê« "
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
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   4050
      Width           =   2490
   End
   Begin VB.ComboBox xYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8625
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   3975
      Width           =   3165
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
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   36
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
         Height          =   330
         Left            =   1440
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   615
         Width           =   540
      End
      Begin VB.CheckBox xMeeting 
         Appearance      =   0  'Flat
         Caption         =   "€—«„… Ã„⁄Ì… ⁄„Ê„Ì… "
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
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   315
         Width           =   1965
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
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "€—«„… »⁄œ :"
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
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   645
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "»‰Êœ «Œ—Ì"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   540
      Width           =   3930
      Begin MSDBCtls.DBCombo xLocker 
         Bindings        =   "items6.frx":0000
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   270
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DBCombo1"
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
      Begin MSDBCtls.DBCombo xYacht 
         Bindings        =   "items6.frx":0014
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   645
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DBCombo1"
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «··‰‘ :"
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
         Index           =   3
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   675
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·œÊ·«» :"
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
         Index           =   2
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   300
         Width           =   1020
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
      Left            =   1500
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   375
      Visible         =   0   'False
      Width           =   1140
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   900
      Top             =   75
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
      Height          =   1905
      Left            =   6165
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2025
      Width           =   5640
      Begin VB.ComboBox xSex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "items6.frx":0028
         Left            =   1440
         List            =   "items6.frx":0035
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1350
         Width           =   2940
      End
      Begin VB.TextBox xAge2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   1485
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   945
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
         TabIndex        =   5
         Top             =   945
         Width           =   555
      End
      Begin VB.CheckBox xIsMember 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "«·⁄÷Ê ‰›”Â :"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1350
      End
      Begin MSDBCtls.DBCombo xRelation 
         Bindings        =   "items6.frx":0046
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Top             =   585
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label9 
         Caption         =   "«·‰Ê⁄ :"
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
         TabIndex        =   27
         Top             =   1425
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·‰ ”‰ :"
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
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ ”‰ :"
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»‰œ «·ﬁ—«»… :"
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
         TabIndex        =   20
         Top             =   675
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «·»‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6165
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   540
      Width           =   5640
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   1125
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   675
         Width           =   2940
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3510
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         Height          =   330
         Left            =   3510
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·»‰œ :"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·»‰œ :"
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
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·»‰œ :"
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
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1125
         Width           =   810
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   825
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
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11880
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   11880
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… «·»‰Êœ"
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
         Left            =   10650
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   75
         Width           =   1140
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
         Height          =   390
         Left            =   4350
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "≈÷«›…"
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
         Left            =   8550
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
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
         Height          =   390
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   90
         Width           =   990
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
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
         Left            =   7500
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Õ–›"
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
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
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
         Left            =   6450
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1710
      Width           =   2670
      Begin VB.CheckBox xBasicOld 
         Appearance      =   0  'Flat
         Caption         =   "√”«”Ì ··⁄÷Ê «·ﬁœÌ„"
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
         TabIndex        =   41
         Top             =   1800
         Width           =   1920
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
         TabIndex        =   35
         Top             =   600
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   900
         Width           =   705
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   7290
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   135
         Width           =   915
      End
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   4050
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "FrmPaidItem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As Recordset, Deletetable1 As Recordset
Dim oSearch As New Search10
Dim SectionItemTable As Recordset
Dim RecordCountTable As Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
cmdInform.Enabled = (nMode = LoadMode)
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
xDesca.Text = ""
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
xCode.Text = CardTable.code
xDesca.Text = TurnValue(CardTable.Desca, Null, "")
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
Sub MyReplace()
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
CardTable!code = xCode.Text
CardTable!Desca = TurnValue(xDesca, "", Null)
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
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·»‰œ €Ì— „”Ã·"
    Exit Function
End If

If xDesca.Text = "" Then
    MsgBox "»Ì«‰ «·»‰œ €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
CardTable.MoveLast
xCode.Text = CardTable.code + 1
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
Private Sub CmdSave_Click()
'msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub

'If Not MsgBox(msgBoxStr, 4) = 6 Then
'    CmdUndo_Click
'    Exit Sub
'End If
MyReplace
Inform " „  «·Õ›Ÿ »‰Ã«Õ"
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.FindFirst "code = " & xCode.Text
    myload
End If

End Sub
Private Sub CmdUndo_Click()
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
    For nrow = 2 To grid1.rows - 1
        For nCol = 4 To grid1.Cols - 1
            mydb.Execute "INSERT INTO FILE2_40(ITEM,[YEAR],[TYPE],SECTION,[VALUE],DISCOUNT,BASIC) " & _
                        " VALUES ( " & _
                        addvalue(xCode.Text) & "," & _
                        addvalue(xYear.ItemData(nYear)) & "," & _
                        addvalue(.TextMatrix(0, nCol)) & "," & _
                        addstring(.TextMatrix(nrow, 0)) & "," & _
                        addvalue(.TextMatrix(nrow, 2)) & "," & _
                        addvalue(.TextMatrix(nrow, 3)) & "," & _
                        IIf(Val(.TextMatrix(nrow, nCol)) = 0, "FALSE", "TRUE") & ")"
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
                 addstring(printTable!code) & "," & _
                 addstring(Chr(254) & Replace(printTable!Desca & "", " ", Chr(254) & " " & Chr(254))) & "," & _
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
PrintGrdNew.doprint grid1, 1.5, 0, "ÿ»«⁄… »‰œ : " & xDesca.Text, "„Ê”„ " & xYear.Text, , , False, False, 10
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

DATA1.DatabaseName = MdbPath
DATA1.RecordSource = "Select Code,DescA from FILE0_00 WHERE FLAG = 0 order by Code"
xRelation.BoundColumn = "Code"
xRelation.ListField = "DESCA"

DATA2.DatabaseName = MdbPath
DATA2.RecordSource = "Select Code,DescA from file0_00 Where Flag = 6"
xLocker.BoundColumn = "Code"
xLocker.ListField = "DESCA"

DATA3.DatabaseName = MdbPath
DATA3.RecordSource = "Select Code,DescA from file0_00 Where Flag = 8"
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
    grid1.TextMatrix(0, i) = paidTypeTable.code
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
    .TextMatrix(.rows - 1, 0) = SectionTable!code
    .TextMatrix(.rows - 1, 1) = SectionTable!Desca
    SectionTable.MoveNext
Loop Until SectionTable.EOF
End With
End Sub
Private Sub LoadGrd()
Set GRDTABLE = mydb.OpenRecordset("Select * From File2_40 " & _
               " where Item = " & xCode.Text & _
               " and [year] = " & xYear.ItemData(xYear.ListIndex))

With grid1
For nrow = 2 To grid1.rows - 1
    .TextMatrix(nrow, 2) = ""
    .TextMatrix(nrow, 3) = ""
     For nCol = 4 To grid1.Cols - 1
         .TextMatrix(nrow, nCol) = ""
     Next
Next

If GRDTABLE.RecordCount = 0 Then
    Exit Sub
End If

For nrow = 2 To grid1.rows - 1
    For nCol = 4 To grid1.Cols - 1
        GRDTABLE.FindFirst "Section = " & MyParn(grid1.TextMatrix(nrow, 0)) & _
                           " AND [TYPE] = " & grid1.TextMatrix(0, nCol)
        If Not GRDTABLE.NoMatch Then
            .TextMatrix(nrow, 2) = TurnValue(GRDTABLE!Value, Null, "")
            .TextMatrix(nrow, 3) = TurnValue(GRDTABLE!Discount, Null, "")
            .TextMatrix(nrow, nCol) = IIf(GRDTABLE!basic, -1, 0)
        End If
    Next
Next
End With
End Sub
Private Sub myreplaceGrd()
With grid1
mydb.Execute "DELETE * FROM FILE2_40 WHERE YEAR = " & xYear.ItemData(xYear.ListIndex) & _
             " AND ITEM = " & xCode.Text

For nrow = 2 To grid1.rows - 1
    For nCol = 4 To grid1.Cols - 1
        mydb.Execute "INSERT INTO FILE2_40(ITEM,[YEAR],[TYPE],SECTION,[VALUE],DISCOUNT,BASIC) " & _
                    " VALUES ( " & _
                    addvalue(xCode.Text) & "," & _
                    addvalue(xYear.ItemData(xYear.ListIndex)) & "," & _
                    addvalue(.TextMatrix(0, nCol)) & "," & _
                    addstring(.TextMatrix(nrow, 0)) & "," & _
                    addvalue(.TextMatrix(nrow, 2)) & "," & _
                    addvalue(.TextMatrix(nrow, 3)) & "," & _
                    IIf(Val(.TextMatrix(nrow, nCol)) = 0, "FALSE", "TRUE") & ")"
    Next
Next
End With
End Sub
Private Sub xYear_Click()
If xCode.Text = "" Then Exit Sub
LblYear.Caption = xYear.ItemData(xYear.ListIndex)
End Sub
