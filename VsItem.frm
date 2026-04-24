VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTItem 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   10365
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
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1275
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1470
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
      Left            =   2625
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1320
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
         Top             =   960
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
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1410
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
      Begin VB.Label Label2 
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
         Left            =   4800
         TabIndex        =   4
         Top             =   210
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
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
      Bindings        =   "VsItem.frx":0000
      Height          =   7560
      Left            =   150
      TabIndex        =   9
      Top             =   2025
      Width           =   14865
      _cx             =   26220
      _cy             =   13335
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
      Left            =   0
      Top             =   450
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
      Left            =   0
      Top             =   600
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
End
Attribute VB_Name = "VsTItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï √—’œ… „»Ì⁄«  „‘ —Ì«  ··√’‰«›  "
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    MyLoad
End Sub
Private Sub Form_Load()
    xDate1.Text = "3-6-2009"
    xDate2.Text = Format(Date, "dd-mm-yyyy")
    
'    CSTRSAL = "select max(date) as m_date , item from file6_20 group by item "
'    If LastSalTable.State = adStateOpen Then LastSalTable.Close
'    LastSalTable.Open CSTRSAL, CON, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    data1.ConnectionString = CON.ConnectionString
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = CON.ConnectionString
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
    Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    
    data3.ConnectionString = CON.ConnectionString
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    
    Set grid1.DataSource = data4
    data4.ConnectionString = CON.ConnectionString
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub MyLoad()

cWhere = " date < " & DateSq(xDate1.Text)
cField4 = myiif(cWhere, "[IN] - [OUT]") & " AS F_BAL"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '2' OR TYPE = '7' )"
cField5 = myiif(cWhere, "[IN] - [OUT]") & " AS T_PURCH"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField6 = myiif(cWhere, "[OUT] - [IN]") & " AS T_SALES"


cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField7 = myiif(cWhere, "Val(( FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))") & " AS TV_SALES"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField8 = myiif(cWhere, "Val((FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_10.PRICE & '') ") & " AS TV_PRICE"

cWhere = " date <= " & DateSq(xDate2.Text)
cField11 = myiif(cWhere, "[IN] - [OUT]") & " AS endbal "
With grid1
'                           0               1                 2                3
    cStrAll = "  select file1_10.item , file1_10.desca , FILE1_10.PRICE , FILE1_10.PRICE2 , " & _
                cField4 & " , " & cField5 & " , " & cField6 & " , " & cField7 & " , " & cField8 & ",  '  ' AS N9   , '  ' AS N10   , " & cField11 & _
                " from ( FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM ) inner join file1_50 on file1_50.code = file1_10.group  WHERE TRUE "
    If xGroup.BoundText <> "" Then cStrAll = cStrAll & " AND [file1_10.GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStrAll = cStrAll & " AND file1_50.group   = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStrAll = cStrAll & " AND [Section] = " & xSection.BoundText
    cStrAll = cStrAll & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PRICE , FILE1_10.PRICE2 "
    data4.RecordSource = cStrAll
    data4.Refresh
    
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 14
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "”⁄— Ã„·…"
    .TextMatrix(0, 3) = "”⁄— ﬁÿ«⁄Ï"
    .TextMatrix(0, 4) = "—’Ìœ " & Format(xDate1.Text, "dd-mm-yyyy")
    .TextMatrix(0, 5) = "„‘ —Ì« "
    .TextMatrix(0, 6) = "„»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄«  ›⁄·Ì…"
    .TextMatrix(0, 8) = "ﬁÌ„… „»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 9) = "ﬁÌ„… Œ’„ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·Œ’„"
    .TextMatrix(0, 11) = "—’Ìœ " & Format(xDate2.Text, "dd-mm-yyyy")
    .TextMatrix(0, 12) = "‰”»… «·»Ì⁄"
    .TextMatrix(0, 13) = "√Œ—  «—ÌŒ »Ì⁄"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 700
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    .ColWidth(7) = 900
    .ColWidth(8) = 900
    .ColWidth(9) = 900
    .ColWidth(10) = 900
    .ColWidth(11) = 800
    .ColWidth(12) = 700
    .ColWidth(13) = 0
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
'   .ColDataType(12) = flexDTDouble
'   .ColDataType(13) = flexDTDate
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For I = 1 To .Rows - 1
        .TextMatrix(I, 9) = Format(Val(.TextMatrix(I, 8)) - Val(.TextMatrix(I, 7)), "#0.00")
        If Val(.TextMatrix(I, 9)) <> 0 And Val(.TextMatrix(I, 8)) <> 0 Then .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 9)) / Val(.TextMatrix(I, 8)) * 100, "#0.00")
    
        If Val(.TextMatrix(I, 6)) <> 0 Then
            If (Val(.TextMatrix(I, 4)) + Val(.TextMatrix(I, 5))) <> 0 Then .TextMatrix(I, 12) = Format(Val(.TextMatrix(I, 6)) / (Val(.TextMatrix(I, 4)) + Val(.TextMatrix(I, 5))) * 100, "#0.00")
        End If
        
'        LastSalTable.Filter = "item = " & MyParn(.TextMatrix(I, 0))
'        If Not LastSalTable.EOF Then .TextMatrix(I, 13) = Format(LastSalTable!m_date, "dd-mm-yyyy")
    Next I
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
    
    If Val(.TextMatrix(1, 9)) <> 0 And Val(.TextMatrix(1, 8)) <> 0 Then .TextMatrix(1, 10) = Format(Val(.TextMatrix(1, 9)) / Val(.TextMatrix(1, 8)) * 100, "#0.00")
    If Val(.TextMatrix(1, 6)) <> 0 Then
        If (Val(.TextMatrix(1, 4)) + Val(.TextMatrix(1, 5))) <> 0 Then .TextMatrix(1, 12) = Format(Val(.TextMatrix(1, 6)) / (Val(.TextMatrix(1, 4)) + Val(.TextMatrix(1, 5))) * 100, "#0.00")
    End If
    If .Rows > 1 Then .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
    End With
End Sub
Private Sub grid1_dblClick()
    If grid1.Col <= 5 Then
        Load StoreMove
        StoreMove.XITEM.Text = grid1.TextMatrix(grid1.Row, 0)
        StoreMove.Show 1
    Else
        ShowSalItem.Show 1
    End If
End Sub
