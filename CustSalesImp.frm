VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CustSalesImp 
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
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   11085
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xCode 
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
         Height          =   315
         Left            =   8355
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox XDATE2 
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox xDoc_No 
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
         Left            =   7125
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Bindings        =   "CustSalesImp.frx":0000
         DataSource      =   "DATA2"
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   975
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XFACTNAME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   3585
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
         TabIndex        =   15
         Top             =   1005
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·—”«·…"
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
         Left            =   5880
         TabIndex        =   13
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   4560
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   690
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” ‰œ"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   960
      End
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   1020
      Width           =   1200
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
      Bindings        =   "CustSalesImp.frx":0014
      Height          =   8010
      Left            =   150
      TabIndex        =   4
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
Attribute VB_Name = "CustSalesImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SalesTable As New ADODB.Recordset
Dim ChqTable As New ADODB.Recordset
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „»Ì⁄«  ··⁄„·«¡ „‰ —”«·… —ﬁ„ " & XCODE.Text & " » «—ÌŒ " & xDate1.Text
    
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
    myload
    FixGrid
End Sub
Private Sub Form_Load()
    
    xdate2.Text = Format(Date, "dd-mm-yyyy")
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "SELECT * FROM FILE3_50 "
    Set xGrCust.RowSource = DATA2
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
    Dim cStrSal As String
   '                    0               1               2                                                                       3       4           5           6
    cStrSal = " SELECT FILE3_10.CODE, FILE3_10.DESCA, Sum([T_SAL_RET].[T_PRICE]) AS T_PRICE, Sum(T_SAL_RET.TOTAL) AS T_SALES, ' ' AS N4, ' ' AS N5, Sum([T_SAL_RET].[T_COST]) AS TCOST, ' ' AS N7, ' ' AS N8, ' ' AS n9 " & _
                " FROM (FILE1_10 INNER JOIN file1_50 ON FILE1_10.[GROUP] = file1_50.CODE) INNER JOIN (FILE3_10 INNER JOIN T_SAL_RET ON FILE3_10.CODE = T_SAL_RET.code) ON FILE1_10.ITEM = T_SAL_RET.ITEM  " & _
                " WHERE FILE1_10.ITEM IN (SELECT ITEM FROM FILE7_60 INNER JOIN FILE7_60H ON FILE7_60H.DOC_NO = FILE7_60.DOC_NO   WHERE FILE7_60H.DOC_NO  = " & MyParn(xDoc_no.Text) & " ) "
    If xGrCust.BoundText <> "" Then cStrSal = cStrSal & " and file3_10.[GROUP] = " & MyParn(xGrCust.BoundText)
    cStrSal = cStrSal & " GROUP BY FILE3_10.CODE, FILE3_10.DESCA, FILE3_10.DESCA ORDER BY FILE3_10.CODE  "
    data1.RecordSource = cStrSal
    data1.Refresh
    FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 10
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    .TextMatrix(0, 2) = "„»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄«  ›⁄·Ì…"
    .TextMatrix(0, 4) = "≈Ã„«·Ï Œ’„"
    .TextMatrix(0, 5) = "‰”»… Œ’„"
    .TextMatrix(0, 6) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 7) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 8) = "‰”»… —»Õ „»Ì⁄« "
    .TextMatrix(0, 9) = "‰”»… „»Ì⁄«  «·⁄„Ì·"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1100
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble

    .ColFormat(2) = "#0.00"
    .ColFormat(3) = "#0.00"
    .ColFormat(4) = "#0.00"
    .ColFormat(5) = "#0.00"
    .ColFormat(6) = "#0.00"
    .ColFormat(7) = "#0.00"
    .ColFormat(8) = "#0.00"
    .ColFormat(9) = "#0.00"
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .Subtotal flexSTSum, -1, 2, "#0", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0", vbYellow, vbRed, True, "  "
    
    If .Rows > 0 Then
        For i = 1 To .Rows - 1
            .TextMatrix(i, 4) = Format(Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 3)), "#0.00")
            If Val(.TextMatrix(i, 2)) <> 0 Then .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(i, 2)) * 100, "#0.00")
            .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i, 6)), "#0.00")
            If Val(.TextMatrix(i, 3)) <> 0 Then .TextMatrix(i, 8) = Format(Val(.TextMatrix(i, 7)) / Val(.TextMatrix(i, 3)) * 100, "#0.00")
            If Val(.TextMatrix(1, 3)) > 0 Then .TextMatrix(i, 9) = Format(Val(.TextMatrix(i, 3)) / Val(.TextMatrix(1, 3)) * 100, "#0.00")
        Next i
        .SubtotalPosition = flexSTAbove
        
        .Subtotal flexSTSum, -1, 4, "#0", vbYellow, vbRed, True, "  "
        .Subtotal flexSTSum, -1, 6, "#0", vbYellow, vbRed, True, "  "
        .Subtotal flexSTSum, -1, 7, "#0", vbYellow, vbRed, True, "  "
        If .Rows > 1 Then
            If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 5) = Format(Val(.TextMatrix(1, 4)) / Val(.TextMatrix(1, 2)) * 100, "#0.00")
            If Val(.TextMatrix(1, 3)) <> 0 Then .TextMatrix(1, 8) = Format(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 3)) * 100, "#0.00")
        End If
    End If
    End With
End Sub
Private Sub grid1_dblClick()
    Load VsTCustSales
    VsTCustSales.XCODE.Text = grid1.TextMatrix(grid1.Row, 0)
    VsTCustSales.xCustName.Caption = grid1.TextMatrix(grid1.Row, 1)
    VsTCustSales.Show
End Sub

Private Sub xDOC_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(3, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "SELECT FILE7_60H.DOC_NO, file4_10.DESCA, FILE7_60H.date, FILE7_60H.FACTNAME FROM file4_10 RIGHT JOIN FILE7_60H ON file4_10.CODE = FILE7_60H.code "
    Generalarray(2) = "Order by FILE7_60H.DATE DESC "
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·„Ê—œ-  «·—”«·…"
    listarray(0, 1) = "(%%DESCA%% OR %%FACTNAME%%)"
    
    GrdArray(0, 0) = "—ﬁ„ „” ‰œ"
    GrdArray(0, 1) = 1500
    
    GrdArray(1, 0) = "«·«”„"
    GrdArray(1, 1) = 3000
    
    GrdArray(2, 0) = " «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    GrdArray(3, 0) = "«·—”«·…"
    GrdArray(3, 1) = 2000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub
Sub myProc()
    xDoc_no.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xDate1.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 2)
    xCodeDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    XFACTNAME.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 3)
    Unload Search3
Exit Sub
myerror:
Unload Search
End Sub

