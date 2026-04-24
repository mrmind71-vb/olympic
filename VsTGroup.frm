VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTGroup 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   9015
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   4920
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "VsTGroup.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsTGroup.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsTGroup.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "VsTGroup.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   9405
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   5730
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1590
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   195
         TabIndex        =   5
         Top             =   1275
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
         Left            =   195
         TabIndex        =   6
         Top             =   915
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
         Left            =   195
         TabIndex        =   7
         Top             =   555
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
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1365
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
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1605
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
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   570
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
         Left            =   3750
         TabIndex        =   4
         Top             =   225
         Width           =   675
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8685
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6675
      Left            =   90
      TabIndex        =   16
      Top             =   1845
      Width           =   15045
      _cx             =   26538
      _cy             =   11774
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
Attribute VB_Name = "VsTGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub CmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï √—’œ… „»Ì⁄«  „‘ —Ì«  ··„Ã„Ê⁄«   "
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    PrintGrd.doprint grid1, 0.9, -2, cHead1, cHead2, , False, False, 9, , Array(1)
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
    
    
    Set grid1.DataSource = data4
    data4.ConnectionString = strCon
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cwhere As String
If IsDate(xDate1.Text) Then cwhere = " date < " & DateSq(xDate1.Text)

cField1 = myiif(cwhere, "[IN] - [OUT]") & " AS F_BAL"

cwhere = ""

If IsDate(xDate1.Text) Then cwhere = " date >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(xDate2.Text)

cField2 = myiif(cwhere & turn(cwhere, " And ") & " ( TYPE = '2' OR TYPE = '7' )", "[IN] - [OUT]") & " AS T_PURCH"

cField3 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "[OUT] - [IN]") & " AS T_SALES"

cField4 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "(FILE1_11.OUT - FILE1_11.[IN])* FILE1_11.PRICE * (1-(FILE1_11.DISCOUNT/100))") & " AS TV_SALES"

cField5 = myiif(cwhere & turn(cwhere, " And ") & "(TYPE = '6' OR TYPE = '3')", "(FILE1_11.OUT - FILE1_11.[IN] )* (FILE1_11.PRICE_ORG) ") & " AS TV_PRICE"

cwhere = ""
If IsDate(xDate2.Text) Then cwhere = " date <= " & DateSq(xDate2.Text)
cField6 = myiif(cwhere, "[IN] - [OUT]") & " AS endbal "

With grid1
'                           0                           1                           2                               3                           4                           5
    cString = "  SELECT file1_10sc.code as c_sec, file1_10sc.desca  as secdesca ,  FILE1_50G.code as mgrcode , FILE1_50G.DESCA as mgrdesca ,  FILE1_50.code as grcode , FILE1_50.DESCA as grdesc ,  " & _
                cField1 & " , " & cField2 & " , " & cField3 & " , " & cField6 & " , " & cField4 & " , " & cField5 & _
                " FROM (((FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_10SC ON FILE1_10.[SECTION] =  FILE1_10SC.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE"
    If xGroup.BoundText <> "" Then cString = cString & turn(cString) & " FILE1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cString = cString & turn(cString) & "  file1_50.[Group]  = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cString = cString & turn(cString) & " [Section] = " & xSection.BoundText
    cString = cString & " GROUP BY FILE1_50.code, FILE1_50.DESCA, FILE1_50G.code, FILE1_50G.DESCA, FILE1_10SC.CODE, FILE1_10SC.DESCA ORDER BY FILE1_10SC.DESCA , FILE1_50G.DESCA , FILE1_50.DESCA "
    data4.RecordSource = cString
    data4.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 14
    .RowHeight(0) = 1000
    .WordWrap = True
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "ﬁ”„"
    .TextMatrix(0, 2) = "ﬂÊœ"
    .TextMatrix(0, 3) = "„Ã„Ê⁄… —∆Ì”Ì…"
    .TextMatrix(0, 4) = "ﬂÊœ"
    .TextMatrix(0, 5) = "„Ã„Ê⁄…"
    
    .TextMatrix(0, 6) = "—’Ìœ √Ê·"
    .TextMatrix(0, 7) = "„‘ —Ì« "
    .TextMatrix(0, 8) = "„»Ì⁄« "
    .TextMatrix(0, 9) = "—’Ìœ"
    
    .TextMatrix(0, 10) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 11) = "ﬁÌ„… „»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 12) = "ﬁÌ„… Œ’„ „»Ì⁄« "
    .TextMatrix(0, 13) = "‰”»… «·Œ’„"
    
        
    .ColHidden(0) = True
    .ColWidth(1) = 2200
    .ColWidth(2) = 0
    .ColWidth(3) = 2200
    .ColWidth(4) = 0
    .ColWidth(5) = 2200
    
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 800
    .ColHidden(11) = True
    .ColHidden(12) = True
    .ColHidden(13) = True
    
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For i = 1 To .Rows - 1
        .TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 11)) - Val(.TextMatrix(i, 10)), "#0.00")
        If Val(.TextMatrix(i, 12)) <> 0 And Val(.TextMatrix(i, 11)) <> 0 Then .TextMatrix(i, 13) = Format(Val(.TextMatrix(i, 12)) / Val(.TextMatrix(i, 11)) * 100, "#0.00")

    Next i
    .SubtotalPosition = flexSTAbove
        
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 10, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 12, "#0", vbRed, vbYellow, True, "  "
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
        If Val(.TextMatrix(1, 12)) <> 0 And Val(.TextMatrix(1, 11)) <> 0 Then .TextMatrix(1, 13) = Format(Val(.TextMatrix(1, 12)) / Val(.TextMatrix(1, 11)) * 100, "#0.00")
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
        .TextMatrix(1, 2) = "«·≈Ã„«·Ï"
        .TextMatrix(1, 3) = "«·≈Ã„«·Ï"
        .TextMatrix(1, 4) = "«·≈Ã„«·Ï"
        .TextMatrix(1, 5) = "«·≈Ã„«·Ï"
        .MergeCells = flexMergeFree
        .MergeRow(1) = True
    End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub Grid1_DblClick()
    With grid1
    If .Row > 1 Then
        Load grditem1
        grditem1.xGroup.BoundText = .TextMatrix(.Row, 4)
        grditem1.xGroupMain.BoundText = .TextMatrix(.Row, 2)
        grditem1.xSection.BoundText = .TextMatrix(.Row, 0)
        grditem1.Show
    End If
    End With
End Sub
