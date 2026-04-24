VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cutfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√„— «·ﬁ’"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   14715
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   600
      Left            =   1575
      TabIndex        =   26
      Top             =   5805
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
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   20
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1020
      Left            =   6885
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   630
      Width           =   7665
      Begin VB.TextBox xResp_Name 
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
         Left            =   2610
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   3525
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "DATE"
         Top             =   225
         Width           =   1320
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
         Left            =   4815
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   135
         TabIndex        =   3
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
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
      Begin VB.Label Label4 
         Caption         =   "≈”„ «·„”∆Ê· :"
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   630
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   630
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1020
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   630
      Width           =   1515
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   180
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   1485
      Top             =   270
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
   Begin VB.Frame Frame8 
      Height          =   615
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1035
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
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
         TabIndex        =   14
         Top             =   180
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
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
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move Last"
         Top             =   180
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   375
      Left            =   2205
      Top             =   720
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
      Height          =   3300
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1620
      Width           =   14370
      Begin VSFlex7Ctl.VSFlexGrid grid20 
         Height          =   900
         Left            =   585
         TabIndex        =   29
         Top             =   630
         Visible         =   0   'False
         Width           =   1230
         _cx             =   2170
         _cy             =   1587
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
         Left            =   11520
         MaxLength       =   15
         TabIndex        =   4
         Top             =   180
         Width           =   1320
      End
      Begin VSFlex7Ctl.VSFlexGrid grid10 
         Height          =   2625
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   14190
         _cx             =   25030
         _cy             =   4630
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   6315
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
         Left            =   12960
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   1155
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   3750
      Left            =   180
      TabIndex        =   27
      Top             =   4950
      Width           =   14370
      _cx             =   25347
      _cy             =   6615
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "Cutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim cStrBox As String
Dim oSearchItem As New Search3
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "Date"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "MOSM"
aInsert(2, 1) = addstring(xMosm.BoundText)

aInsert(3, 0) = "Resp_NAME"
aInsert(3, 1) = addstring(xResp_Name.Text)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("CUTH", "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "CUTH")
Else
    con.Execute CreateUpdate(aInsert, "CUTH", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
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
cString = " DELETE  FROM CUT WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " AND MODEL = " & MyParn(xModel.Text)
con.Execute cString
With grid10
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                cString = "Insert into CUT (doc_no,Item,Model,Quant)" & _
                           "Values(" & _
                           addstring(xDoc_No.Text) & "," & _
                           addstring(grid20.TextMatrix(nRow, nCol)) & "," & _
                           addstring(xModel.Text) & "," & _
                           Val(.TextMatrix(nRow, nCol)) & _
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
    con.Execute "Delete From CUT where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete From CUTH where Doc_No = " & MyParn(xDoc_No.Text)
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
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT CUTH.Doc_No, CONVERT(VARCHAR(10),CUTH.[Date],111),Resp_Name,Mosm" & _
                  " FROM CUTH "
Generalarray(2) = " order by CUTH.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-«”„ «·„ÊœÌ·"
listarray(0, 1) = "##CUTH.Date## " & _
                  " Or (DOC_NO in (Select DOC_NO FROM CUT INNER JOIN FILE1_10H ON CUT.MODEL = FILE1_10H.MODEL WHERE %%FILE1_10H.DESCA%%))"

listarray(1, 0) = "≈”„ «·„”∆Ê·-«·„Ê”„"
listarray(1, 1) = "(%%Resp_Name%% or 'cFilter' Like Mosm)"
                  

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«”„ «·„”∆Ê·"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "«·„Ê”„"
GrdArray(3, 1) = 1000


searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«„— ﬁ’"
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
myload
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
openCon con
cStrBox = StrBox
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM CUTH ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Set Grid2.DataSource = DATA11
DATA11.ConnectionString = strCon

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM MOSM"
Set xMosm.RowSource = data1
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"

fixgrdModel
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
    Fixgrd
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

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    'grid1.TextMatrix(Row, 1) = RetSeek("file5_00", "ndxcode", Trim(grid1.TextMatrix(Row, 0)), "Desca")
    grid1.TextMatrix(Row, 1) = GetDesca("select Desca from file5_00 where code = " & MyParn(grid1.TextMatrix(Row, 0)))
End If
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 1 Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 5 Then
    grid1.Editable = IIf(Val(grid1.TextMatrix(grid1.Row, 6)) = 0, flexEDKbdMouse, flexEDNone)
ElseIf grid1.Col = 6 Then
    grid1.Editable = IIf(Val(grid1.TextMatrix(grid1.Row, 5)) = 0, flexEDKbdMouse, flexEDNone)
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> 0 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If Grid2.TextMatrix(Grid2.Row, 0) <> "" Then
            cString = "DELETE  FROM CUT WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " AND MODEL = " & MyParn(Grid2.TextMatrix(Grid2.Row, 0))
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
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    If grid1.Rows > 3 Then
        grid1.TextMatrix(grid1.Rows - 2, 2) = grid1.TextMatrix(grid1.Rows - 3, 2)
        grid1.TextMatrix(grid1.Rows - 2, 3) = grid1.TextMatrix(grid1.Rows - 3, 3)
    End If
End If
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
xResp_Name.Text = CardTable!RESP_NAME & ""
xMosm.BoundText = CardTable!mosm & ""
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT CUT.ITEM,CUT.QUANT  " & _
               " FROM CUT " & _
               " where Doc_no = " & MyParn(xDoc_No.Text)
    DATA10.RecordSource = cString
    DATA10.Refresh
End With


cString = "SELECT CUT.MODEL,FILE1_10H.DESCA,SUM(Quant)  " & _
           " FROM (CUT INNER JOIN FILE1_10 ON CUT.ITEM = FILE1_10.ITEM) INNER JOIN FILE1_10H ON FILE1_10.MODEL = FILE1_10H.MODEL" & _
           " where CUT.Doc_no = " & MyParn(xDoc_No.Text) & _
           " GROUP BY CUT.MODEL,FILE1_10H.DESCA"
DATA11.RecordSource = cString
DATA11.Refresh

Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Newflag("CUTH", "DOC_NO"))
xMosm.BoundText = ""
xResp_Name.Text = ""
xDate.Text = ""
grid1.Rows = 1
Grid2.Rows = 1
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

Private Sub VSFlexGrid1_Click()

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
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!Code & ";" & boxtable!Desca
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub Fixgrd()
With Grid2
    .Cols = 3
    .FormatString = "ﬂÊœ «·„ÊœÌ·|" & "≈”„ «·„ÊœÌ·|" & "«·ﬂ„Ì…"
    .ColWidth(0) = 1500
    .ColWidth(1) = 6000
    .ColWidth(2) = 1000
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
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
'FixCost cFieldas, cField
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
xDate.BackColor = &HC0FFFF
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.BackColor = &HC0FFFF
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
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub xResp_Name_GotFocus()
xResp_Name.BackColor = &HC0FFFF
xResp_Name.SelStart = 0
xResp_Name.SelLength = Len(xResp_Name.Text)
End Sub

