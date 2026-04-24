VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VsPrice2 
   Caption         =   "„ «»⁄… «”⁄«— «·‘—«¡"
   ClientHeight    =   9705
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   9705
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1140
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1260
      Width           =   2040
      Begin VB.CommandButton CmdUndo 
         Height          =   465
         Left            =   1035
         MaskColor       =   &H00FFFFFF&
         Picture         =   "VsPrice2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "⁄—÷"
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
         Left            =   1035
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ÿ»«⁄…"
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
         TabIndex        =   16
         Top             =   180
         Width           =   915
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
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   585
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   13050
      Begin VB.TextBox xDesca 
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
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1935
         Width           =   3345
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
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1575
         Width           =   1455
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
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1215
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xGroup 
         Bindings        =   "VsPrice2.frx":289A
         DataSource      =   "data3"
         Height          =   315
         Left            =   7905
         TabIndex        =   3
         Top             =   870
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Bindings        =   "VsPrice2.frx":28AE
         DataSource      =   "DATA2"
         Height          =   315
         Left            =   7905
         TabIndex        =   4
         Top             =   510
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Bindings        =   "VsPrice2.frx":28C2
         DataSource      =   "data1"
         Height          =   315
         Left            =   7905
         TabIndex        =   5
         Top             =   150
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Bindings        =   "VsPrice2.frx":28D6
         Height          =   2115
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Width           =   7755
         _cx             =   13679
         _cy             =   3731
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
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
         TabBehavior     =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
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
         Height          =   195
         Index           =   4
         Left            =   11475
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1980
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Height          =   195
         Left            =   11475
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ :"
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
         Height          =   195
         Index           =   0
         Left            =   11475
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1665
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Index           =   1
         Left            =   11460
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11385
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   555
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   3
         Left            =   11460
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   165
         Width           =   1230
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   14400
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1395
      Top             =   -360
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   105
      Top             =   -60
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1395
      Top             =   -60
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   105
      Top             =   -360
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   45
      Top             =   -360
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsPrice2.frx":28EB
      Height          =   6930
      Left            =   90
      TabIndex        =   13
      Top             =   2430
      Width           =   15090
      _cx             =   26617
      _cy             =   12224
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
      TabBehavior     =   0
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
Attribute VB_Name = "VsPrice2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFilesave As String
Dim cString As String
Dim cString2 As String
Dim cFilter1 As String
Dim cFilter As String
Dim osearchitem As New Search3
Dim nInd As Byte
Dim con As New ADODB.Connection

Private Sub Cmd_Print_Click()
Dim cHead1 As String
Dim cHead2 As String
cHead1 = "„ «»⁄… √”⁄«— ‘—«¡ Œ·«· › —… "
cHead2 = BetweenString(Format(xdate1.Text, "YYYY-MM-DD"), Format(XDATE2.Text, "YYYY-MM-DD"))
grid1.ColHidden(0) = True
Load PrintGrd
PrintGrd.doprint grid1, 1, -2, cHead1, cHead2, , False, True, 10
PrintGrd.Show 1
grid1.ColHidden(0) = False
End Sub

Private Sub CmdAdditem_Click()
ItemsLookupAll Me, osearchitem
End Sub

Private Sub cmdDelinv_Click()

End Sub

Private Sub cmdExit_Click()
Unload Me
Set VsPrice2 = Nothing

End Sub
Sub fillgrd()
Dim nSubTotal As Double
Dim nTotal As Double
Dim nT As Double
Dim cSubGroup As String, CGROUP As String
Dim invTable As New ADODB.Recordset
Me.MousePointer = 11

grid1.Sort = flexSortNone
xGroup.Enabled = True
nSubTotal = 0
nTotal = 0
nT = 0
cFilter = ""
cString = "SELECT  FILE1_10.ITEM, FILE1_10.DESCA,Sum(FILE7_20.QUANT) AS SumQUANT,Sum(FILE7_20.TOTAL),FILE1_10.COST, Max(FILE7_20.PRICE) AS MaxPRICE, Min(FILE7_20.PRICE),COALESCE(Max(FILE7_20.PRICE),0) - COALESCE(MIN(FILE7_20.PRICE),0)    " & _
          " FROM ((FILE7_20 INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM) INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO ) inner join file1_50 on file1_10.[GROUP] = file1_50.code"

If xGroup.BoundText <> "" Then cString = cString & turn(cString) & " file1_10.[GROUP]  = " & xGroup.BoundText
If xGroupMain.BoundText <> "" Then cString = cString & turn(cString) & " file1_50.[Group]  = " & xGroupMain.BoundText
If xSection.BoundText <> "" Then cString = cString & turn(cString) & "  [Section] = " & xSection.BoundText
If Trim(xDesca.Text) <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "file1_10.DESCA")

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " file7_20H.Date >= " & DateSq(xdate1.Text)
End If

If IsDate(XDATE2.Text) Then
    cString = cString & turn(cString) & " file7_20H.Date <= " & DateSq(XDATE2.Text)
End If

Dim cOr As String
With grid2
For I = 1 To grid2.Rows - 1
    If Trim(.TextMatrix(I, 0)) <> "" Then
        cOr = cOr & turn(cOr, " OR ") & "File1_10.ITEM = " & MyParn(grid2.TextMatrix(I, 0))
    End If
Next
cOr = turn(cOr, "(") & cOr & turn(cOr, ")")
If cOr <> "" Then
    cString = cString & turn(cString) & cOr
End If
End With
cString = cString & " GROUP BY FILE1_10.ITEM, FILE1_10.DESCA  , FILE1_10.COST  "
invTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Set DATA11.Recordset = invTable
Fixgrd
With grid1
'.Rows = 1

'Do Until invTable.EOF
'    .AddItem ""
'   .TextMatrix(.Rows - 1, 1) = TurnValue(invTable!Item, Null, "")
'   .TextMatrix(.Rows - 1, 2) = TurnValue(invTable!Desca, Null, "")
'   .TextMatrix(.Rows - 1, 3) = Format(invTable!SUMQUANT, "##0.00")
'   .TextMatrix(.Rows - 1, 4) = Format(invTable!Sumtotal, "##0.000")
'   .TextMatrix(.Rows - 1, 5) = invTable!COST & ""
'   .TextMatrix(.Rows - 1, 6) = Format(invTable!MAXPRICE, "##0.00")
'   .TextMatrix(.Rows - 1, 7) = Format(invTable!MINPRICE, "##0.00")
'   .TextMatrix(.Rows - 1, 8) = Format(Val(.TextMatrix(.Rows - 1, 6)) - Val(.TextMatrix(.Rows - 1, 5)), "##0.00")
'   invTable.MoveNext
'Loop
'.Sort = 1
Me.MousePointer = 1
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 3, "##0.00", , RGB(255, 0, 0), True

.Subtotal flexSTSum, -1, 4, "##0.00", , RGB(255, 0, 0), True
.SubtotalPosition = flexSTAbove
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlue
End With
End Sub

Private Sub CmdOk_Click()
fillgrd
End Sub
Private Sub CmdUndo_Click()
grid1.Rows = 1
grid2.Rows = 1
grid2.AddItem ""
DefineText Me
End Sub

Private Sub Command1_Click()

End Sub
Private Sub Form_Load()
    cFilesave = App.Path & "\" & Me.Name & "_gd.grd"
    DATA11.ConnectionString = strCon

    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
'   Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    data1.Refresh
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
'   Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    DATA2.Refresh
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
'   Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    data3.Refresh
    
    xdate1.Text = ""
    XDATE2.Text = ""
    myloadgrd
    Fixgrd
    openCon con
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
grid2.SaveGrid cFilesave, flexFileData
Unload Me
Set VsPrice = Nothing
End Sub
Private Sub Fixgrd()
With grid1
.OutlineBar = flexOutlineBarNone
.ExplorerBar = flexExSortShow

.Cols = 9
.TextMatrix(0, 1) = "ﬂÊœ"
.TextMatrix(0, 2) = "’‰›"
.TextMatrix(0, 3) = "Ã. ‘—«¡"
.TextMatrix(0, 4) = "Ã. ﬁÌ„…"
.TextMatrix(0, 5) = "”⁄—  ﬂ·›…"
.TextMatrix(0, 6) = "√⁄·Ï ”⁄—"
.TextMatrix(0, 7) = "√ﬁ· ”⁄—"
.TextMatrix(0, 8) = "«·›—ﬁ"

.RowHeight(0) = 600
.WordWrap = True
.ColWidth(0) = 400
.ColWidth(1) = 1500
.ColWidth(2) = 2500
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100

.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.FixedCols = 1
End With
End Sub
Private Sub grid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookupAll Me, osearchitem
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
With grid2
If grid2.Row = grid2.Rows - 1 Then
    MyAddItem
End If
End With
End Sub
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If (Not validRow(OldRow, True, True)) And OldRow <> grid2.Rows - 1 And OldRow <> 0 And grid2.TextMatrix(OldRow, grid2.Cols - 1) = "" Then
    grid2.RemoveItem OldRow
End If
End Sub
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 Then
    grid2.RemoveItem grid2.Row
    grid2.Select grid2.Rows - 1, 1
    grid2.ShowCell grid2.Rows - 1, 1
End If
End Sub

Private Sub grid2_Validate(Cancel As Boolean)
If (Not validRow(grid2.Row, True, True)) And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "" Then
    grid2.RemoveItem grid2.Row
End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Sub myProc()
Dim nFound As Long
nFound = grid2.FindRow(osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0), , 0)
If nFound > 0 Then
    MsgBox "«·’‰› „ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ " & nFound
    Exit Sub
End If
Dim bNew As Boolean
bNew = grid2.Row = grid2.Rows - 1
grid2.TextMatrix(grid2.Row, 0) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 0)
grid2.TextMatrix(grid2.Row, 1) = osearchitem.grid1.TextMatrix(osearchitem.grid1.Row, 1)
grid2_AfterEdit grid2.Row, 0
If Not bNew Then
    Unload osearchitem
Else
    grid2.Row = grid2.Rows - 1
End If
End Sub
Private Sub MyAddItem()
grid2.AddItem ""
grid2.ShowCell grid2.Rows - 1, 1
End Sub
Private Sub myloadgrd()
With grid2
.TextMatrix(0, 1) = "«·’‰›"
.ColHidden(0) = True
.ColComboList(1) = "..."
.ColWidth(1) = 6000
.ColAlignment(1) = flexAlignRightCenter
Dim fs As New FileSystemObject
If fs.FileExists(cFilesave) Then
    On Error Resume Next
    grid2.LoadGrid cFilesave, flexFileData
    Err.Clear
End If
If grid2.Rows = 1 Then grid2.AddItem ""
grid2.ShowCell grid2.Rows - 1, 1
grid2.Select grid2.Rows - 1, 1
End With
End Sub
