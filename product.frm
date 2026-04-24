VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form productfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«‰ «Ã"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
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
   ScaleHeight     =   10110
   ScaleWidth      =   15090
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid grdtool 
      Height          =   600
      Left            =   1080
      TabIndex        =   38
      Top             =   900
      Visible         =   0   'False
      Width           =   1320
      _cx             =   2328
      _cy             =   1058
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
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
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   600
      Left            =   1080
      TabIndex        =   31
      Top             =   900
      Visible         =   0   'False
      Width           =   1320
      _cx             =   2328
      _cy             =   1058
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   9495
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   5595
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
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   2790
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
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
         Height          =   420
         Left            =   4140
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   675
      Width           =   8610
      Begin VB.TextBox xNotes 
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
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   6900
      End
      Begin VB.TextBox xOrderNo 
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
         Left            =   5670
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1320
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
         Tag             =   "DATE"
         Top             =   180
         Width           =   1770
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
         Left            =   5670
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   5220
         TabIndex        =   4
         Top             =   900
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   900
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
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
         Height          =   240
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Œ“‰ «·ﬁ’ :"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Œ“‰ «·Œ«„«  :"
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   945
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "—ﬁ„ «„— «·ﬁ’ :"
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ :"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   705
      End
      Begin VB.Label Label3 
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
         Height          =   240
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ «·„” ‰œ :"
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1350
      Width           =   1545
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   585
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
         TabIndex        =   11
         Top             =   180
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   450
      Top             =   180
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1755
      Width           =   2040
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
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   465
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
         Height          =   375
         Left            =   555
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   465
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
         Height          =   375
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   465
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
         Height          =   375
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move Last"
         Top             =   180
         Width           =   465
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   375
      Left            =   450
      Top             =   135
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   450
      Top             =   135
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
   Begin VSFlex7Ctl.VSFlexGrid grid20 
      Height          =   540
      Left            =   1080
      TabIndex        =   34
      Top             =   900
      Visible         =   0   'False
      Width           =   1320
      _cx             =   2328
      _cy             =   952
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
      Rows            =   3
      Cols            =   2
      FixedRows       =   3
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
      Editable        =   2
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
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   375
      Left            =   450
      Top             =   135
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Height          =   375
      Left            =   450
      Top             =   135
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      Height          =   3030
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2340
      Width           =   14910
      Begin VB.TextBox xModel 
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
         Left            =   11970
         MaxLength       =   15
         TabIndex        =   7
         Top             =   180
         Width           =   1320
      End
      Begin VSFlex7Ctl.VSFlexGrid grid10 
         Height          =   2355
         Left            =   90
         TabIndex        =   8
         Top             =   585
         Width           =   14730
         _cx             =   25982
         _cy             =   4154
         _ConvInfo       =   1
         Appearance      =   0
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
         Rows            =   3
         Cols            =   2
         FixedRows       =   3
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
         Editable        =   2
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
      Begin VB.Label xModelDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   5505
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
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
         Left            =   13410
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   225
         Width           =   1155
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4605
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5310
      Width           =   14910
      Begin VSFlex7Ctl.VSFlexGrid Grid2 
         Height          =   1005
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   14685
         _cx             =   25903
         _cy             =   1773
         _ConvInfo       =   1
         Appearance      =   0
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
      Begin VSFlex7Ctl.VSFlexGrid Grid3 
         Height          =   3165
         Left            =   135
         TabIndex        =   10
         Top             =   1305
         Width           =   14685
         _cx             =   25903
         _cy             =   5583
         _ConvInfo       =   1
         Appearance      =   0
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
End
Attribute VB_Name = "productfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim oSearchItem As New Search3, osearchitemRaw As New Search3, oSearchOrder As New Search3
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(6, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "Date"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "Mosm"
aInsert(2, 1) = addstring(xMosm.BoundText)

aInsert(3, 0) = "OrderNo"
aInsert(3, 1) = addstring(xOrderNo.Text)

aInsert(4, 0) = "store1"
aInsert(4, 1) = addstring(xStore1.BoundText)

aInsert(5, 0) = "store2"
aInsert(5, 1) = addstring(xStore2.BoundText)

aInsert(6, 0) = "Notes"
aInsert(6, 1) = addstring(xNotes.Text)

'On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE1_90H", "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE1_90H")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_90H", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
myreplaceGrd2
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim cString As String, nRow As Integer, cFile As String
cString = " DELETE  FROM FILE1_90 WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " AND MODEL = " & MyParn(xModel.Text)
con.Execute cString
With grid10
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                cString = "Insert into FILE1_90 (doc_no,Item,Model,Quant,Cost)" & _
                           "Values(" & _
                           addstring(xDoc_No.Text) & "," & _
                           addstring(grid20.TextMatrix(nRow, nCol)) & "," & _
                           addstring(xModel.Text) & "," & _
                           Val(.TextMatrix(nRow, nCol)) & "," & _
                           Val(.TextMatrix(1, nCol)) & _
                           ")"
                con.Execute cString
            End If
        Next
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
ElseIf ActiveControl.Name = xOrderNo.Name Then
    xOrderNo.Text = oSearchOrder.grid1.TextMatrix(oSearchOrder.grid1.Row, 0)
    Unload oSearchOrder
ElseIf ActiveControl.Name = Grid3.Name Then
    nFound = FoundOtheritem(Grid3.Row, Grid3.Col, Trim(osearchitemRaw.grid1.TextMatrix(osearchitemRaw.grid1.Row, 0)))
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    Grid3.TextMatrix(Grid3.Row, 0) = osearchitemRaw.grid1.TextMatrix(osearchitemRaw.grid1.Row, 0)
    Grid3.TextMatrix(Grid3.Row, 2) = "1"
    GrdDesc Grid3.Row
    If Grid3.Row = Grid3.Rows - 1 Then
        Grid3.TextMatrix(Grid3.Rows - 1, 2) = ""
        Grid3.AddItem ""
        Grid3.Select Grid3.Rows - 1, 0
    ElseIf Grid3.Row = Grid3.Rows - 2 Then
        Grid3.TextMatrix(Grid3.Rows - 2, 2) = ""
        Grid3.Select Grid3.Rows - 1, 0
    End If
Else
    ActiveControl.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    xModelDesca.Caption = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 1)
    xModel_LostFocus
    Unload oSearchItem
End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE1_90 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete From FILE1_91 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete From FILE1_90H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.EOF Then CardTable.MoveFirst
        myload
    End If
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
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE1_90H.Doc_No, CONVERT(VARCHAR(10),FILE1_90H.[Date],111),FILE0_40.DESCA,FILE0_40_1.DESCA,NOTES " & _
                  " FROM (FILE1_90H INNER JOIN FILE0_40 ON FILE1_90H.STORE1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_90H.STORE2 = FILE0_40_1.CODE "
Generalarray(2) = " order by FILE1_90H.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-«”„ «·„ÊœÌ·"
listarray(0, 1) = "##FILE1_90H.Date## " & _
                  " Or (DOC_NO in (Select DOC_NO FROM FILE1_90 INNER JOIN FILE1_10H ON FILE1_90.MODEL = FILE1_10H.MODEL WHERE %%FILE1_10H.DESCA%%))"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "„Œ“‰ «·Œ«„« "
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„Œ“‰ «·ﬁ’"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "„·«ÕŸ« "
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ «·«‰ «Ã"
Search3.Show 1
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
xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1, 6)
mydefine
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
xModel.Text = ""
xModelDesca.Caption = ""
myDefineGrdModel
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
If CardTable.EOF And CardTable.BOF Then mydefine Else myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
makeMyValidate Me
openCon con
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM FILE1_90H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Set Grid2.DataSource = DATA11
DATA11.ConnectionString = strCon

Set Grid3.DataSource = DATA12
DATA12.ConnectionString = strCon

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM MOSM"
Set xMosm.RowSource = data1
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE0_40"
Set xStore1.RowSource = DATA2
xStore1.ListField = "DESCA"
xStore1.BoundColumn = "CODE"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE0_40"
Set xStore2.RowSource = DATA2
xStore2.ListField = "DESCA"
xStore2.BoundColumn = "CODE"

fixgrdModel
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
    Fixgrd
    Fixgrd2
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
SetKbLayout Lang_AR
On Error Resume Next
Unload SearchItems
Set SearchItems = Nothing
Err.Clear
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> 0 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If Grid2.TextMatrix(Grid2.Row, 0) <> "" Then
            cString = "DELETE  FROM FILE1_90 WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " AND MODEL = " & MyParn(Grid2.TextMatrix(Grid2.Row, 0))
            con.Execute cString
        End If
        con.CommitTrans
        myDefineGrdModel
        myloadgrd
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub GRID3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, osearchitemRaw
End If

If KeyCode = 46 And grid1.Row <> 0 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If Grid3.TextMatrix(Grid3.Row, 0) <> "" Then
            cString = "DELETE  FROM FILE1_91 WHERE ID = " & Grid3.TextMatrix(Grid3.Row, Grid3.Cols - 1)
            con.Execute cString
        End If
        con.CommitTrans
        Grid3.RemoveItem Grid3.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Function MYVALID() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    MsgBox " «—ÌŒ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "DD-MM-YYYY")
xStore1.BoundText = CardTable!store1 & ""
xStore2.BoundText = CardTable!Store2 & ""
xOrderNo.Text = CardTable!orderno & ""
xNotes.Text = CardTable!Notes & ""
xMosm.BoundText = CardTable!mosm & ""
Handlecontrols LoadMode
myloadgrd
myloadgrd2
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE1_90.ITEM,FILE1_90.QUANT,FILE1_90.COST  " & _
               " FROM FILE1_90 " & _
               " where Doc_no = " & MyParn(xDoc_No.Text)
    DATA10.RecordSource = cString
    DATA10.Refresh
End With


cString = "SELECT FILE1_90.MODEL,FILE1_10H.DESCA,SUM(Quant),SUM(FILE1_90.QUANT * FILE1_90.COST)  " & _
           " FROM FILE1_90  INNER JOIN FILE1_10H ON FILE1_90.MODEL = FILE1_10H.MODEL" & _
           " where FILE1_90.Doc_no = " & MyParn(xDoc_No.Text) & _
           " GROUP BY FILE1_90.MODEL,FILE1_10H.DESCA"
DATA11.RecordSource = cString
DATA11.Refresh
Fixgrd
End Sub
Private Sub myloadgrd2()
cString = "SELECT FILE1_91.ITEM,FILE1_10.DESCA,FILE1_91.QUANT,FILE1_91.COST, FILE1_91.QUANT * FILE1_91.COST ,FILE1_91.ID  " & _
           " FROM FILE1_91  INNER JOIN FILE1_10 ON FILE1_91.ITEM = FILE1_10.ITEM" & _
           " where FILE1_91.Doc_no = " & MyParn(xDoc_No.Text) & _
           " ORDER BY FILE1_91.ROW"
DATA12.RecordSource = cString
DATA12.Refresh
Grid3.AddItem ""
Fixgrd2
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Newflag("FILE1_90H", "DOC_NO"))
xDate.Text = ""
xOrderNo.Text = ""
xMosm.Text = ""
xStore1.BoundText = ""
xStore2.BoundText = ""
xNotes.Text = ""
grid1.Rows = 1
Grid2.Rows = 1
Grid3.Rows = 1
Grid3.AddItem ""

xModel.Text = ""
xModelDesca.Caption = ""
myDefineGrdModel
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = (nMode = LoadMode And bEdit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
cmdsave.Enabled = bEdit
CmdDelInv.Enabled = bEdit
End Sub
Private Sub Grid2_Click()
If Trim(Grid2.TextMatrix(NewRow, 0)) <> Trim(xModel.Text) Then
    xModel.Text = Trim(Grid2.TextMatrix(Grid2.Row, 0))
    xModelDesca.Caption = Grid2.TextMatrix(Grid2.Row, 1)
    myDefineGrdModel
    myloadGrdModel
End If
End Sub
Private Sub Grid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Grid2_Click
End Sub

Private Sub xDoc_No_LostFocus()
xDoc_No.BackColor = &H80000005
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub Fixgrd()
With Grid2
    .Cols = 4
    .FormatString = "ﬂÊœ «·„ÊœÌ·|" & "≈”„ «·„ÊœÌ·|" & "«·ﬂ„Ì…|" & "«· ﬂ·›…"
    .ColWidth(0) = 1500
    .ColWidth(1) = 6000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub Fixgrd2()
With Grid3
    .Cols = 6
    .FormatString = "ﬂÊœ «·Œ«„…|" & "≈”„ «·Œ«„…|" & "«·ﬂ„Ì…|" & "«· ﬂ·›…|" & "«·«Ã„«·Ì|"
    .ColWidth(0) = 1500
    .ColWidth(1) = 6000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColHidden(.Cols - 1) = True
End With
End Sub
Private Sub myloadGrdModel()
Dim aret As Variant, cFieldas As String, cField As String

myDefineGrdModel

aret = retFields
If aret(0) = "" Then Exit Sub

cField = aret(0)
cFieldas = aret(1)

FillItem cFieldas, cField
FixCost cFieldas, cField, "cost", 1
fixgrdModel
End Sub
Private Function retFields()
Dim aret(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(xModel.Text) & " group by SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!Scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!Scal & "]" & " as " & "[" & FieldTable!Scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!Scal & "]"
    End If
    FieldTable.MoveNext
Loop
aret(0) = cField
aret(1) = cFieldas
retFields = aret
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub myDefineGrdModel()
grid20.Rows = 3
grid20.Cols = 2

grid10.Rows = 3
grid10.Cols = 2

grid10.MergeCells = flexMergeRestrictRows
grid10.TextMatrix(0, 1) = "«·„ﬁ«”"
grid10.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid10.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
'grid1.FixedRows = 3
End Sub

Private Sub xModel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ModelLookupAll Me, oSearchItem
End If
End Sub
Private Sub xModel_LostFocus()
xModel.BackColor = &H80000005
myDefineGrdModel
xModelDesca.Caption = ""
If Trim(xModel.Text) = "" Then Exit Sub
xModelDesca.Caption = GetDesca("Select DESCA from FILE1_10H WHERE MODEL = " & MyParn(xModel.Text))
myDefineGrdModel
myloadGrdModel
SetKbLayout Lang_AR
End Sub
Private Sub FillItem(cFieldas, cField)
Dim GRDTABLE As New ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item,col_color from file1_10 WHERE MODEL = " & MyParn(xModel.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.col_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid10.Cols = GRDTABLE.Fields.Count: grid20.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grid10.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
Next

Do Until GRDTABLE.EOF
    grid20.AddItem ""
    grid10.AddItem ""
    For nCol = 0 To GRDTABLE.Fields.Count - 1
        If nCol <= 1 Then
            grid10.TextMatrix(grid20.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
        Else
            grid20.TextMatrix(grid20.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
            nFoundRow = grid1.FindRow(GRDTABLE.Fields(nCol).Value & "", , 0)
            If nFoundRow <> -1 Then
                 grid10.TextMatrix(grid10.Rows - 1, nCol) = grid1.TextMatrix(nFoundRow, 1)
             Else
                 grid10.TextMatrix(grid10.Rows - 1, nCol) = ""
             End If
        End If
    Next
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub fixgrdModel()
grid10.ColWidth(0) = 500
grid10.ColWidth(1) = 1200
nColWidth = (grid10.Width - 200 - grid10.ColWidth(0) - grid10.ColWidth(1)) / grid10.Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1000 Then nColWidth = 1000
For nCol = 2 To grid10.Cols - 1
    grid10.ColWidth(nCol) = nColWidth
    grid10.ColAlignment(nCol) = flexAlignCenterCenter
Next
grid10.ColHidden(0) = True
grid10.RowHidden(1) = True
grid10.RowHidden(2) = True
grid10.TextMatrix(0, 1) = "«·„ﬁ«”"
End Sub
Private Sub FixCost(cFieldas, cField, Optional cFieldAdd As String = "Cost", Optional nRow As Integer = 1)
Dim cString As String
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal," & cFieldAdd & " from file1_10 WHERE MODEL = " & MyParn(xModel.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(" & cFieldAdd & ")" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdTextdu
If Not loctable.EOF Then
    For nCol = 2 To grid10.Cols - 1
        grid10.TextMatrix(nRow, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Sub xdate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xModel_GotFocus()
xModel.BackColor = &HC0FFFF
xModel.SelStart = 0
xModel.SelLength = Len(xModel.Text)
SetKbLayout Lang_EN
End Sub
Private Sub xDate_LostFocus()
xDate.BackColor = &H80000005
End Sub
Private Sub xMosm_LostFocus()
xMosm.BackColor = &H80000005
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub

Private Sub xOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then OrderLookupAll Me, oSearchOrder
End Sub

Private Sub xStore1_LostFocus()
xStore1.BackColor = &H80000005
If Not xStore1.MatchedWithList Then xStore1.BoundText = ""
End Sub
Private Sub xStore2_LostFocus()
xStore2.BackColor = &H80000005
If Not xStore2.MatchedWithList Then xStore1.BoundText = ""
End Sub
Private Sub xNotes_GotFocus()
xNotes.SelStart = 0
xNotes.SelLength = Len(xNotes.Text)
xNotes.BackColor = &HC0FFFF
End Sub
Private Sub xOrderNo_GotFocus()
xOrderNo.SelStart = 0
xOrderNo.SelLength = Len(xOrderNo.Text)
xOrderNo.BackColor = &HC0FFFF
End Sub
Private Sub xMosm_GotFocus()
xMosm.BackColor = &HC0FFFF
End Sub
Private Sub xStore1_GotFocus()
xStore1.BackColor = &HC0FFFF
End Sub
Private Sub xStore2_GotFocus()
xStore2.BackColor = &HC0FFFF
End Sub
Private Sub xNotes_LostFocus()
xNotes.BackColor = &H80000005
End Sub
Private Sub xOrderNo_LostFocus()
xOrderNo.BackColor = &H80000005
End Sub
Private Sub myreplaceGrd2()
Dim aInsert(4, 1)
With Grid3
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 2))

        aInsert(3, 0) = "cost"
        aInsert(3, 1) = Val(.TextMatrix(i, 3))

        aInsert(4, 0) = "row"
        aInsert(4, 1) = i
        
        If Grid3.TextMatrix(i, Grid3.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "File1_91")
        Else
            con.Execute CreateUpdate(aInsert, "File1_91", " where ID = " & Grid3.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub GRID3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Grid3.Col = 0 Then GrdDesc Grid3.Row
CalcTotals
End Sub
Private Sub GRID3_EnterCell()
If Grid3.Col = 0 Or Grid3.Col = 2 Or Grid3.Col = 3 Then
    Grid3.Editable = flexEDKbdMouse
Else
    Grid3.Editable = flexEDNone
End If
End Sub
Private Sub GRID3_GotFocus()
With Grid3
    If Grid3.Row = 0 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub
Private Sub GRID3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And Grid3.Row <> Grid3.Rows - 1 Then Grid3.AddItem "", Grid3.Row
End Sub
Private Sub GRID3_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Grid3.Row = Grid3.Rows - 1 Then Grid3.AddItem ""
End Sub
Private Sub GRID3_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(Grid3.EditText) <> "" Then
    cItem = GetDesca("select item from file1_10 where item = " & MyParn(Grid3.EditText)) & ""
    If cItem = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
        Exit Sub
    End If
    
    nFound = FoundOtheritem(Row, Col, Trim(Grid3.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Cancel = True
    End If
End If
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To Grid3.Rows - 2
    If i <> nRow Then
        If Trim(Grid3.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub GrdDesc(ByVal Row As Long)
Grid3.TextMatrix(Row, 1) = ""
If Grid3.TextMatrix(Row, 0) = "" Then Exit Sub

aret = aGetDesca("select desca,Cost from file1_10 where item = " & MyParn(Grid3.TextMatrix(Grid3.Row, 0)))
If UBound(aret) > 0 Then
    Grid3.TextMatrix(Row, 1) = aret(1) & ""
    Grid3.TextMatrix(Row, 3) = aret(2) & ""
End If
End Sub
Private Function CalcTotals()
With Grid3
    For i = 1 To .Rows - 2
        .TextMatrix(i, 4) = Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 3))
    Next
End With
End Function
Private Sub xDate_Validate(Cancel As Boolean)
With xDate
If (Not IsDate(.Text)) And Trim(.Text) <> "" Then .Text = ""
.Text = Format(.Text, "dd-mm-yyyy")
End With
End Sub
Private Sub add_barcode_Click()
If Not MYVALID Then Exit Sub
Dim tBarCode As New ADODB.Recordset
If grid1.Rows = 1 Then Exit Sub
tBarCode.Open "addprint", con, adOpenKeyset, adLockReadOnly, adCmdTable

With grid1
con.Execute "DELETE  FROM ADDPRINT WHERE DOC_NO = " & MyParn(xDoc_No.Text)
For i = 1 To grid1.Rows - 2
        con.Execute "Insert Into ADDPRINT(Doc_no,code,Item,Quant,isPrint) " & _
                " Values(" & _
                addstring(xDoc_No.Text) & "," & _
                addstring(xcode.Text) & "," & _
                addstring(.TextMatrix(i, 1)) & "," & _
                Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 7)) & "," & _
                "1" & _
                ")"
Next
End With
End Sub
