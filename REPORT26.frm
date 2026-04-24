VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm26 
   Caption         =   "»Ì«‰«  «·«⁄÷«¡"
   ClientHeight    =   3030
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   10080
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdExel 
      Height          =   555
      Left            =   135
      Picture         =   "REPORT26.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "⁄—÷"
      Top             =   2025
      Width           =   2220
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1665
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   1440
      Width           =   1500
   End
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3195
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   180
      Width           =   2760
      Begin VB.CheckBox xSafeOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
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
         Height          =   390
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   810
         Width           =   2400
      End
      Begin VB.CheckBox xDiedOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «·„ Ê›ÌÌ‰ ›ﬁÿ"
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
         Height          =   390
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   135
         Width           =   1950
      End
      Begin VB.CheckBox xDropOnly 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
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
         Height          =   390
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   450
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   180
      Width           =   2400
      Begin VB.CheckBox xSafe 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
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
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   765
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin VB.CheckBox xDrop 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
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
         Height          =   435
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   405
         Width           =   2040
      End
      Begin VB.CheckBox xDied 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «·„ Ê›ÌÌ‰"
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
         Height          =   345
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   135
         Value           =   1  'Checked
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   180
      Width           =   4515
      Begin VB.TextBox xcode1 
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox xCode2 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1680
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   900
         Width           =   1410
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   900
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   540
         Width           =   3120
         _ExtentX        =   5503
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
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   1260
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   1665
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   2070
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "„‰ —ﬁ„"
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
         Index           =   7
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "·„ Ì”œœ"
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
         Index           =   3
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "«Œ— ”œ«œ ›Ì"
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
         Index           =   2
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "”œœ „Ê”„"
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
         Index           =   4
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1305
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4815
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   26
      Top             =   2790
      Visible         =   0   'False
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4470
      Left            =   -13005
      TabIndex        =   27
      Top             =   4050
      Visible         =   0   'False
      Width           =   13560
      _cx             =   23918
      _cy             =   7885
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      Cols            =   3
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSCommand cmdPdf 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   2385
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2025
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   979
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
      Picture         =   "REPORT26.frx":27EB
      Caption         =   "Pdf ÿ»«⁄…"
      ButtonStyle     =   1
      PictureAlignment=   10
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "REPORT26.frx":4DB6
   End
End
Attribute VB_Name = "reportfrm26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Private Sub cmdApply_Click()
doprint
End Sub

Private Sub cmdExel_Click()
Dim cString As String
cString = doprint(False, True)

If cString = "" Then Exit Sub

Set grid1.DataSource = data10
Set data10.Recordset = mySet(cString, con)

With grid1
.ColWidth(0) = 800
.ColWidth(1) = 2000
.ColWidth(2) = 1000
.ColWidth(3) = 3000
.ColWidth(4) = 1700
.ColWidth(5) = 1700
.ColWidth(6) = 1500

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«”„ «·‰«œÌ"
.TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
.TextMatrix(0, 3) = "«·«”„ «·—»«⁄Ì"
.TextMatrix(0, 4) = "«·—ﬁ„ «·ﬁÊ„Ì"
.TextMatrix(0, 5) = "«·ÊŸÌ›…"
.TextMatrix(0, 6) = "«· ·Ì›Ê‰ «·„Õ„Ê·"
.ColHidden(2) = True
End With

ToFileExel2 grid1, , , , , 1, , , , , , Me
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint(Optional pPdf As Boolean = False, Optional bExcel As Boolean = False) As String
Dim aHeader(12)
Dim cString As String
cString = "Select ROW_NUMBER() OVER(ORDER BY FILE1_10.CODE ASC) AS Row_Number ,'‰«œÌ «·«Ê·Ì„»Ì «·”ﬂ‰œ—Ì' as club_desca,file1_10.CODE,FILE1_10.DESCA,FILE1_10.ID_NO,JOB_CODES.DESCA AS Job_Desca,FILE1_10.MOBIL" & _
          " From File1_10 LEFT JOIN JOB_CODES ON FILE1_10.JOB = JOB_CODES.CODE"

If IsNumeric(cmdYear(0).Tag) Then
    aHeader(0) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If

If ValidNum(xcode1.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code " & IIf(ValidNum(xCode2.text), " >= ", " = ") & addvalue(xcode1.text)
    aHeader(1) = IIf(ValidNum(xCode2.text), BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : "), "—ﬁ„ ⁄÷ÊÌ… :" & xcode1.text)
End If

If ValidNum(xCode2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code <= " & addvalue(xCode2.text)
    aHeader(1) = BetweenString(xcode1.text, xCode2.text, "„‰ —ﬁ„ ⁄÷ÊÌ… : ", "Õ Ì —ﬁ„ ⁄÷ÊÌ… : ")
End If


If IsNumeric(cmdYear(1).Tag) Then
    aHeader(2) = "«Œ— ”œ«œ ·Â„ ›Ì „Ê”„ " & cmdYear(1).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) = " & cmdYear(1).Tag
End If

If IsNumeric(cmdYear(2).Tag) Then
    aHeader(3) = "·„ Ì”œœÊ« „Ê”„ " & cmdYear(2).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) < " & cmdYear(2).Tag
End If

If xType.MatchedWithList Then
    aHeader(4) = "›∆… «·⁄÷ÊÌ… : " & xType.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.TYPE = " & addvalue(xType.BoundText)
End If


If IsDate(xDate1.text) Then
    aHeader(5) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1)"
End If

If IsDate(xDate2.text) Then
    aHeader(5) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE <= " & DateSq(xDate2.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If xSafeOnly.Value = 0 Then
    If xSafe.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
        aHeader(6) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
    Else
        'aHeader(7) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
    End If
End If

If xDropOnly.Value = 0 Then
    If xDrop.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
        'aHeader(8) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    Else
        aHeader(7) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    End If
End If


If xDropOnly.Value = 1 Then
    aHeader(8) = "«·”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "[DROP] = 1"
End If

If xSafeOnly.Value = 1 Then
    aHeader(9) = "Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 1)"
End If


If xDiedOnly.Value = 1 Then
    aHeader(10) = "«·„ Ê›ÌÌ‰ ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "Died = 1"
ElseIf xDiedOnly.Value = 0 Then
    If xDied.Value = 0 Then
       cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
        aHeader(10) = "»œÊ‰ «·„ Ê›ÌÌ‰"
    Else
        'aHeader(12) = "⁄—÷ «·„ Ê›ÌÌ‰"
    End If
End If


If cWhere2 <> "" Then
    cString = cString & " WHERE " & cWhere2
ElseIf cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If bExcel Then
    doprint = cString & " ORDER BY FILE1_10.CODE"
    Exit Function
End If


Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset

Me.MousePointer = 11
Set sourcetable = myCmd(cString, con)

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = sourcetable!code
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(sourcetable!ID_NO)
    temptable!str4 = TurnValue(sourcetable!JOB_desca)
    temptable!str5 = sourcetable!Mobil
    temptable!str7 = sourcetable!club_desca
        
    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    If pPdf Then
        FixPrinter Report1, 1
        Report1.Destination = crptToPrinter
    End If
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT26.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
    Me.MousePointer = 0
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function
Private Function fillgrd()
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset
Dim aHeader(12)
cString = "Select file1_10.*,TYPE_CODES.DescA as type_Desca,dbo.f_last_year_date(FILE1_10.CODE) AS DATE_LAST " & _
          " From File1_10 left join TYPE_codes on File1_10.TYPE = TYPE_codes.Code"

If IsNumeric(cmdYear(0).Tag) Then
    aHeader(0) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) >= " & cmdYear(0).Tag
End If


If IsNumeric(cmdYear(1).Tag) Then
    aHeader(1) = "«Œ— ”œ«œ ·Â„ ›Ì „Ê”„ " & cmdYear(1).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) = " & cmdYear(1).Tag
End If

If IsNumeric(cmdYear(2).Tag) Then
    aHeader(2) = "·„ Ì”œœÊ« „Ê”„ " & cmdYear(2).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(FILE1_10.CODE) < " & cmdYear(2).Tag
End If

If xType.MatchedWithList Then
    aHeader(3) = "›∆… «·⁄÷ÊÌ… : " & xType.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.TYPE = " & addvalue(xType.BoundText)
End If

If Trim(xNotes.text) <> "" Then
    aHeader(4) = "«·„·ÕÊŸ… : " & xNotes.text
    cWhere = cWhere & turn(cWhere, " and ") & MyParnAnd(xNotes.text, "file1_10.notes")
End If

Dim myWhere As String
If Val(xYears.text) > 0 And IsNumeric(cmdYear(0).Tag) Then
    myWhere = myWhere & "FILE6_20H.YEARS_DESCA LIKE '" & getYears(Val(xYears.text), Val(cmdYear(0).Tag)) & "%'"
    myWhere = myWhere & turn(myWhere, " and ") & "FILE6_20H.DATE >= " & DateSq(GetField("SELECT DATE1 FROM YEARS_CODES WHERE CODE = " & cmdYear(0).Tag, con))
    myWhere = myWhere & turn(myWhere, " and ") & "FILE6_20H.DATE <= " & DateSq(GetField("SELECT DATE2 FROM YEARS_CODES WHERE CODE = " & cmdYear(0).Tag, con))
    myWhere = myWhere & turn(myWhere, " and ") & "(NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1"
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE " & myWhere & ")"
End If

If IsDate(xDate1.text) Then
    aHeader(5) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE >= " & DateSq(xDate1.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1)"
End If

If IsDate(xDate2.text) Then
    aHeader(6) = "„”œœ " & BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT FILE6_20H.CODE FROM FILE6_20H INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE WHERE DATE <= " & DateSq(xDate2.text) & " AND (NOT FORM_NO IS NULL) AND PAID_TYPES.IS_PAID = 1 )"
End If

If xSafeOnly.Value = 0 Then
    If xSafe.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 0)"
        aHeader(7) = "»œÊ‰ Õ«›ŸÌ «·⁄÷ÊÌ…"
    Else
        aHeader(7) = "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ…"
    End If
End If

If xDropOnly.Value = 0 Then
    If xDrop.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & " (FILE1_10.[DROP] = 0)"
        aHeader(8) = "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    Else
        aHeader(8) = "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
    End If
End If


If xDropOnly.Value = 1 Then
    aHeader(9) = "«·”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "[DROP] = 1"
End If

If xSafeOnly.Value = 1 Then
    aHeader(10) = "Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & " (dbo.f_save(FILE1_10.CODE) = 1)"
End If

If Check1.Value Then
    cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_date(FILE1_10.CODE) IS NULL"
    aHeader(11) = "·Ì” ··⁄÷Ê  «—ÌŒ ”œ«œ"
End If

If xDiedOnly.Value = 1 Then
    aHeader(12) = "«·„ Ê›ÌÌ‰ ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & "Died = 1"
ElseIf xDiedOnly.Value = 0 Then
    If xDied.Value = 0 Then
       cWhere2 = cWhere & turn(cWhere, " and ") & " (Died = 0)"
        aHeader(12) = "»œÊ‰ «·„ Ê›ÌÌ‰"
    Else
        aHeader(12) = "⁄—÷ «·„ Ê›ÌÌ‰"
    End If
End If


If cWhere2 <> "" Then
    cString = cString & " WHERE " & cWhere2
ElseIf cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If
Me.MousePointer = 11
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = sourcetable!code
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(ArbStr(sourcetable!Address))
    temptable!str4 = sourcetable!Mobil
    temptable!str5 = sourcetable!Phone
    temptable!str6 = TurnValue(ArbStr(myFormat_p(sourcetable!DATE_LAST)))

    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    'If xInstall.Value = 1 Then
    '    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT2_1.rpt"
    'Else
        Report1.ReportFileName = sPath_App & "\REPORTS\REPORT2.rpt"
    'End If
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
    Me.MousePointer = 0
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function

Private Sub cmdPdf_Click()
doprint True
End Sub

Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub

Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM TYPE_CODES ORDER BY CODE", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Code"

FixRpImage Me
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub

Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(0).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Function getYears(nYears As Long, nSeason As Long)
Dim loctable As New ADODB.Recordset
loctable.Open "select * from Years_codes where code <= " & nSeason & " order by code desc", con, adOpenStatic, adCmdText
Do Until loctable.EOF Or nCount > nYears
    getYears = loctable!code & turn(getYears, ",") & getYears
    loctable.MoveNext
    nCount = nCount + 1
Loop
End Function
Private Sub xCode1_GotFocus()
myGotFocus xcode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xcode1
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub

