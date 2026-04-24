VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTBalCustGR 
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1387
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1020
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
      Top             =   1020
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1440
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBalCustGR.frx":0000
      Height          =   8010
      Left            =   135
      TabIndex        =   9
      Top             =   1575
      Width           =   14865
      _cx             =   26220
      _cy             =   14129
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
      RowHeightMin    =   500
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   150
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
End
Attribute VB_Name = "VsTBalCustGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „Êﬁ› „Ã„Ê⁄«  «·⁄„·«¡ "
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    cHead2 = " „‰  «—ÌŒ " & Format(xdate1.Text, "YYYY-MM-DD") & " ≈·Ï  «—ÌŒ " & Format(XDATE2.Text, "YYYY-MM-DD")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, cHead1, cHead2, , False, True, 10
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
    data1.RecordSource = "SELECT * FROM FILE3_50 "
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
openCon con
If IsDate(xdate1.Text) Then cwhere = " date < " & DateSq(xdate1.Text)
cField2 = myiif(cwhere, "[SAL] - [PAY]") & " AS F_BAL"

cwhere = ""

If IsDate(xdate1.Text) Then cwhere = " date >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)

cField3 = myiif(cwhere & turn(cwhere, " And ") & " (TYPE = '4')", "[SAL]") & " AS T_SALES"
cField4 = myiif(cwhere & turn(cwhere, " And ") & " (TYPE = '5')", "[PAY]") & " AS T_RET"
cField9 = myiif(cwhere & turn(cwhere, " And ") & " (TYPE <> '4' AND TYPE <> '5' AND TYPE <> 'C' AND TYPE <> 'A' )", "[SAL] - [PAY]") & " AS T_CASH"
cField10 = myiif(cwhere & turn(cwhere, " And ") & " (TYPE <> 'A')", "[PAY]") & " AS T_CHQ"
cField11 = myiif(cwhere & turn(cwhere, " And ") & " (TYPE <> 'C')", "[SAL]") & " AS T_RCHQ"

cwhere = ""
If IsDate(XDATE2.Text) Then cwhere = "date <= " & DateSq(xdate1.Text)
cField12 = myiif(cwhere, "[sal]- [pay] ") & " AS C_BAL"

With grid1
'                           0                           1
    cString = "  select FILE3_50.CODE AS CODE , FILE3_50.DESCA AS DESCA , " & _
                cField2 & " , " & cField3 & " , " & cField4 & " ,   ' 'AS N5 , ' ' AS N6 , ' '  AS N7 , ' ' AS N8 , " & _
                cField9 & " , " & cField10 & " , " & cField11 & " , " & cField12 & " , ' ' AS N13 " & _
                " FROM (FILE3_11 LEFT JOIN FILE3_10 ON FILE3_11.CODE = FILE3_10.CODE) LEFT JOIN file3_50 ON FILE3_10.[GROUP] = file3_50.CODE   WHERE (NOT (FILE3_10.CODE IS NULL)) "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " file3_10.[GROUP]  = " & MyParn(xGrCust.BoundText)
    
    cString = cString & " GROUP BY FILE3_50.CODE , FILE3_50.DESCA  "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FixGrid
fillgrd
End Sub
Sub FixGrid()
    With grid1
    .FrozenCols = 2
    .Cols = 15
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "„Ã„Ê⁄…"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xdate1.Text
    .TextMatrix(0, 3) = "Ã. „»Ì⁄« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 6) = "ﬁÌ„… „»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 7) = "≈Ã„«·Ï ﬁÌ„… «·Œ’„"
    .TextMatrix(0, 8) = "‰”»… «·Œ’„"
    
    .TextMatrix(0, 9) = "”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 10) = "”œ«œ √.ﬁ "
    .TextMatrix(0, 11) = "√.ﬁ  „— œ…"
    .TextMatrix(0, 12) = "—’Ìœ " & XDATE2.Text
    .TextMatrix(0, 13) = "√.ﬁ  Õ  «· Õ’Ì·"
    .TextMatrix(0, 14) = "‰”»… «·„Ã„Ê⁄…"
    
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 900
    .ColWidth(3) = 1300
    .ColWidth(4) = 1300
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 800
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 1000
    .ColWidth(12) = 1300
    .ColWidth(13) = 1300
    .ColWidth(14) = 1000
    
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
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
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    .ColDataType(14) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    End With
End Sub
Sub fillgrd()
Dim SalesTable As New ADODB.Recordset
Dim ChqTable As New ADODB.Recordset
Dim cStrSal As String, cStrChq As String

cStrSal = "SELECT FILE3_10.[GROUP] , Sum(T_SAL_RET.TOTAL) AS T_TOTAL, Sum(T_SAL_RET.T_PRICE) AS T_PRICE , Sum(T_SAL_RET.T_COST) AS T_COST FROM T_SAL_RET INNER JOIN FILE3_10 ON FILE3_10.CODE = T_SAL_RET.CODE  WHERE (NOT(T_SAL_RET.CODE IS NULL)) "
If IsDate(xdate1.Text) Then cStrSal = cStrSal & turn(cStrSal) & "DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cStrSal = cStrSal & turn(cStrSal) & "DATE <= " & DateSq(XDATE2.Text)
cStrSal = cStrSal & " GROUP BY FILE3_10.[GROUP] "
If SalesTable.State = adStateOpen Then SalesTable.Close
SalesTable.Open cStrSal, con, adOpenKeyset, adLockOptimistic, adCmdText

cStrChq = "SELECT FILE3_10.[GROUP] , Sum(VALUE ) AS T_VALUE FROM FILE5_20 INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE5_20.CODE1 WHERE CLOSED = '0' "
cStrChq = cStrChq & " GROUP BY FILE3_10.[GROUP]"
If ChqTable.State = adStateOpen Then ChqTable.Close
ChqTable.Open cStrChq, con, adOpenKeyset, adLockOptimistic, adCmdText

    With grid1
    For I = 1 To .Rows - 1
        .TextMatrix(I, 5) = Val(.TextMatrix(I, 3)) - Val(.TextMatrix(I, 4))
        SalesTable.Find " GROUP = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
        If Not SalesTable.EOF Then
            .TextMatrix(I, 6) = Format(SalesTable!T_PRICE, "#0.00")
            .TextMatrix(I, 7) = Format(Val(.TextMatrix(I, 6)) - Val(.TextMatrix(I, 5)), "FIXED")
            If Val(.TextMatrix(I, 6)) <> 0 Then .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) / Val(.TextMatrix(I, 6)) * 100, "FIXED")
            ChqTable.Find " GROUP = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
            If Not ChqTable.EOF Then
                .TextMatrix(I, 13) = Format(ChqTable!T_VALUE, "FIXED")
            End If
        End If
        For nCol = 2 To 13
            .TextMatrix(I, nCol) = Format(.TextMatrix(I, nCol), "#0.00")
        Next nCol
    Next I
    .SubtotalPosition = flexSTAbove
    For I = 2 To 13
        .Subtotal flexSTSum, -1, I, "#0.00", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 1 Then
        If Val(.TextMatrix(1, 6)) <> 0 Then .TextMatrix(1, 8) = Format(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 6)) * 100, "FIXED")
    End If
    For I = 2 To .Rows - 1
        .TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 5)) / Val(.TextMatrix(1, 5)) * 100, "#0.00")
    Next I
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
    SalesTable.Close
    ChqTable.Close
    Set SalesTable = Nothing
    Set ChqTable = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub Grid1_DblClick()
    If grid1.Row > 1 Then
        Load VsTBalCust
        VsTBalCust.xGrCust.BoundText = grid1.TextMatrix(grid1.Row, 0)
        VsTBalCust.xdate1.Text = xdate1.Text
        VsTBalCust.XDATE2.Text = XDATE2.Text
        VsTBalCust.Show
    End If
End Sub
