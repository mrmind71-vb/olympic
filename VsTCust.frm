VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTCust 
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  «·⁄„·«¡"
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
      Top             =   1095
      Width           =   1200
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1357
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1095
      Width           =   1200
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
      Top             =   1095
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   622
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
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1020
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5010
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   622
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
         Left            =   9555
         TabIndex        =   13
         Top             =   690
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1050
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
      Left            =   525
      Top             =   75
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
      Bindings        =   "VsTCust.frx":0000
      Height          =   7935
      Left            =   150
      TabIndex        =   9
      Top             =   1650
      Width           =   14865
      _cx             =   26220
      _cy             =   13996
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   525
      Top             =   375
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
End
Attribute VB_Name = "VsTCust"
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
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „»Ì⁄«  ·√’‰«›  "
    If xCustName.Caption <> "" Then cHead1 = cHead1 & xCustName.Caption
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    
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
    
    
    DATA5.ConnectionString = CON.ConnectionString
    DATA5.RecordSource = "SELECT * FROM FILE3_50 "
    Set xGrCust.RowSource = DATA5
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set grid1.DataSource = data4
    data4.ConnectionString = CON.ConnectionString
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub MyLoad()

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '6'  )"
cField5 = myiif(cWhere, "[OUT] ") & " AS Q_SALES"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text) & " AND ( TYPE = '3'  )"
cField6 = myiif(cWhere, "[IN ] ") & " AS Q_RET"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text)
cField7 = myiif(cWhere, "[OUT]- [IN] ") & " AS Q_NET"

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text)
cField8 = myiif(cWhere, "Val(( FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_11.PRICE & '')") & " AS TV_8 "

cWhere = " date >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSql(xDate2.Text)
cField9 = myiif(cWhere, "Val(( FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))") & " AS TV_9"






With grid1
'                           0               1                 2                3                4
    cStrAll = "  select FILE1_10.CODE , file3_10.DESCA , SUM(Q_SAL) AS Q1 , SUM(Q_RET) AS Q2 , SUM(S_SAL - Q_RET ) AS Q3 , " & _
              "SUM(TPRICE) AS TP , SUM(DISC_ITEM) AS TD1  , SUM(DISCOUNT) AS DISC , SUM(TOTAL) AS TT "
                " FROM ((FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) LEFT JOIN file1_50 ON FILE1_10.GROUP = file1_50.CODE) LEFT JOIN FILE3_10 ON FILE1_11.CODECUST = FILE3_10.CODE  WHERE ( FILE1_11.TYPE = '3' OR  FILE1_11.TYPE = '6' ) "
    If xGroup.BoundText <> "" Then cStrAll = cStrAll & " AND [file1_10.GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStrAll = cStrAll & " AND file1_50.group   = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStrAll = cStrAll & " AND [Section] = " & xSection.BoundText
    If Xcode.Text <> "" Then cStrAll = cStrAll & " AND [CODECUST] = " & MyParn(Xcode.Text)
    If xGrCust.BoundText <> "" Then cStrAll = cStrAll & " AND FILE3_10.GROUP = " & MyParn(xGrCust.BoundText)
    If xDesca.Text <> "" Then cStrAll = cStrAll & " AND file1_10.DESCA LIKE ('%" & xDesca.Text & "%')   "
    
    cStrAll = cStrAll & " GROUP BY FILE1_50.DESCA , FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PRICE , FILE1_10.PRICE2 "
    data4.RecordSource = cStrAll
    data4.Refresh
    
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 13
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    
    .TextMatrix(0, 2) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 3) = "⁄œœ „— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï „»Ì⁄« "
    
    .TextMatrix(0, 5) = "ﬁÌ„… „»Ì⁄«  ﬁ»· «·Œ’„"
    .TextMatrix(0, 6) = "Œ’„ √’‰«›"
    .TextMatrix(0, 7) = "Œ’„ ›Ê« Ì—"
    .TextMatrix(0, 8) = "’«›Ï ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 9) = "—’Ìœ «·⁄„Ì·"
    
    .TextMatrix(0, 10) = "‰”»… «·Œ’„"
    .TextMatrix(0, 11) = "√Œ—  «—ÌŒ »Ì⁄"
    .TextMatrix(0, 12) = "‰”»… «·⁄„Ì·"
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1200
    .ColWidth(12) = 1000
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    .ColDataType(11) = flexDTDate
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(12) = flexDTDouble

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
'  For I = 1 To .Rows - 1
'       .TextMatrix(I, 10) = Format(Val(.TextMatrix(I, 8)) - Val(.TextMatrix(I, 9)), "#0.00")
'       If Val(.TextMatrix(I, 10)) <> 0 And Val(.TextMatrix(I, 8)) <> 0 Then .TextMatrix(I, 11) = Format(Val(.TextMatrix(I, 10)) / Val(.TextMatrix(I, 8)) * 100, "#0.00")
'
'   Next I
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0.00", vbRed, vbYellow, True, "  "
    
    .Subtotal flexSTSum, -1, 6, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0.00", vbRed, vbYellow, True, "  "
    
'    If .Rows > 1 Then
'        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
'        If Val(.TextMatrix(1, 10)) <> 0 And Val(.TextMatrix(1, 8)) <> 0 Then .TextMatrix(1, 11) = Format(Val(.TextMatrix(1, 10)) / Val(.TextMatrix(1, 8)) * 100, "#0.00")
'    End If
    End With
End Sub
Private Sub grid1_DblClick()
    ShowSalItemCust.Show 1
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If Xcode.Text = "" Then Exit Sub
Xcode.Text = RetZero(Xcode.Text, 6)
xCustName.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(Xcode.Text)) & ""
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
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
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
