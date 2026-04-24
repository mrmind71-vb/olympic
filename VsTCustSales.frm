VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTCustSales 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15045
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
   ScaleHeight     =   10365
   ScaleWidth      =   15045
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1275
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
      Height          =   420
      Left            =   2250
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   900
      Width           =   1275
   End
   Begin VB.CommandButton cmdGo 
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
      Height          =   420
      Left            =   2250
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   450
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11355
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   590
         Width           =   4440
      End
      Begin VB.TextBox Xcode 
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   955
         Width           =   1815
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5985
         TabIndex        =   10
         Top             =   1350
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   982
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   6000
         TabIndex        =   12
         Top             =   611
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   1125
         TabIndex        =   16
         Top             =   1350
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰›"
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
         Left            =   4680
         TabIndex        =   22
         Top             =   640
         Width           =   1065
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   2505
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ ⁄„Ì·"
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
         Index           =   1
         Left            =   4680
         TabIndex        =   19
         Top             =   995
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡"
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
         Index           =   0
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…:"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1350
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1005
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   9555
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -135
      Top             =   300
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Bindings        =   "VsTCustSales.frx":0000
      Height          =   7335
      Left            =   45
      TabIndex        =   9
      Top             =   1800
      Width           =   14865
      _cx             =   26220
      _cy             =   12938
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
      BackColorSel    =   14220542
      ForeColorSel    =   64
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
      SelectionMode   =   1
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -360
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -1170
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
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   -1125
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
End
Attribute VB_Name = "VsTCustSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim oSearch As New Search3
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „»Ì⁄«  ·√’‰«›  "
    If xCustName.Caption <> "" Then cHead1 = cHead1 & xCustName.Caption
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    
    cHead2 = " „‰  «—ÌŒ " & Format(xdate1.Text, "YYYY-MM-DD") & " ≈·Ï  «—ÌŒ " & Format(XDATE2.Text, "YYYY-MM-DD")
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 0.9, -2, cHead1, cHead2, , False, True, 8
    PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    xdate1.Text = "3-6-2009"
    XDATE2.Text = Format(Date, "YYYY-MM-DD")
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
    Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    
    data5.ConnectionString = strCon
    data5.RecordSource = "SELECT * FROM FILE3_50"
    Set xGrCust.RowSource = data5
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set grid1.DataSource = data4
    data4.ConnectionString = strCon
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cwhere As String
If IsDate(xdate1.Text) Then cwhere = " date >= " & DateSq(xdate1.Text)

cField5 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '6' and [OUT]  > 0)", "[OUT]") & " AS salesQuant"
cField6 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '6' AND  [OUT] < 0)", " [OUT]") & " AS retSalesQuant"
cField7 = myiif(cwhere & turn(cwhere, " and ") & "(TYPE = '6')", "OUT") & " AS SalesQuantNet"
cField8 = myiif(cwhere & turn(cwhere, " and ") & "(TYPE = '6')", "OUT * FILE1_11.PRICE") & " AS SalesValue"
cField9 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '6' )", "[DATE]", "''", "MAX") & " AS LastDate"
cField10 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '6' )", "[OUT]*FILE1_11.COST") & " AS SalesCost"
cField11 = myiif("", "[IN] - [OUT]") & " AS LastBalance"


With grid1
'                           0                           1                 2                3                4
    cString = "  select FILE1_50.DESCA AS GRDESCA  , file1_10.item , file1_10.desca , FILE1_10.PRICE3 , " & _
                cField5 & " , " & cField6 & " , " & cField7 & " , " & cField8 & " , " & cField9 & "," & cField10 & ", '  ' AS PROFIT,'' AS PROFITRATE ," & cField11 & _
                " FROM ((FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) LEFT JOIN file1_50 ON FILE1_10.[GROUP] = file1_50.CODE) LEFT JOIN FILE3_10 ON FILE1_11.CODECUST = FILE3_10.CODE  "
    
    If xGroup.BoundText <> "" Then cString = cString & " AND file1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cString = cString & " AND file1_50.[group]   = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cString = cString & " AND [Section] = " & xSection.BoundText
    If XCODE.Text <> "" Then cString = cString & " AND [CODECUST] = " & MyParn(XCODE.Text)
    If xGrCust.BoundText <> "" Then cString = cString & " AND FILE3_10.[GROUP] = " & MyParn(xGrCust.BoundText)
    If xDesca.Text <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "FILE1_10.DESCA")
    If IsDate(XDATE2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
    cString = cString & " GROUP BY file1_50.DESCA, FILE1_10.ITEM, FILE1_10.DESCA, FILE1_10.PRICE3"
    data4.RecordSource = cString
    data4.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 13
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„Ã„Ê⁄…"
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·’‰›"
    .TextMatrix(0, 3) = "Ú”⁄— „” Â·ﬂ"
    
    .TextMatrix(0, 4) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 5) = "⁄œœ „—Ã⁄« "
    .TextMatrix(0, 6) = "’«›Ï „»Ì⁄« "
    .TextMatrix(0, 7) = "’«›Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 8) = "√Œ—  «—ÌŒ »Ì⁄"
    .TextMatrix(0, 9) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 10) = "ﬁÌ„… —»Õ"
    .TextMatrix(0, 11) = "‰”»… —»Õ"
    .TextMatrix(0, 12) = "«·—’Ìœ"
    
    .ColWidth(0) = 1700
    .ColWidth(1) = 1500
    .ColWidth(2) = 2700
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 1200
    .ColWidth(9) = 800
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    
    
'    .ColHidden(3) = True
'    .ColHidden(8) = True
'    .ColHidden(9) = True
'    .ColHidden(10) = True
'    .ColHidden(11) = True
'    .ColHidden(12) = True
'    .ColHidden(13) = True
'    .ColHidden(14) = True
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    .ColDataType(12) = flexDTDate
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
    .ColFormat(11) = "#.#%"
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For I = 1 To .Rows - 1
        .TextMatrix(I, 10) = Val(.TextMatrix(I, 7)) - Val(.TextMatrix(I, 9))
        If Val(.TextMatrix(I, 7)) <> 0 And Val(.TextMatrix(I, 10)) <> 0 Then .TextMatrix(I, 11) = Val(.TextMatrix(I, 10)) / Val(.TextMatrix(I, 7))
    Next I
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 10, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 11, "", vbRed, vbYellow, True, "  "
    
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
        If Val(.TextMatrix(1, 7)) <> 0 And Val(.TextMatrix(1, 10)) <> 0 Then .TextMatrix(1, 11) = Format(Val(.TextMatrix(1, 10)) / Val(.TextMatrix(1, 7)), "#0.00")
    End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub Grid1_DblClick()
    If grid1.Rows > 1 Then ShowSalItemCust.Show 1
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If XCODE.Text = "" Then Exit Sub
XCODE.Text = RetZero(XCODE.Text, 6)
xCustName.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(XCODE.Text)) & ""
End Sub
Sub myProc()
ActiveControl.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xCustName.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„"
oSearch.Show 1
End Sub
