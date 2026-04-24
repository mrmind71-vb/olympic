VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form itemcustold 
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
   Begin VB.Frame Frame1 
      Height          =   1140
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
         Top             =   150
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
      Bindings        =   "itemcustold.frx":0000
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
Attribute VB_Name = "itemcustold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SalesTable As New ADODB.Recordset
Dim ChqTable As New ADODB.Recordset

Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „»Ì⁄«  ··⁄„·«¡ „‰ —”«·… —ﬁ„ " & xCode.Text & " » «—ÌŒ " & xDate1.Text
    
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
    fillgrd

End Sub
Private Sub Form_Load()
    
    XDATE2.Text = Format(Date, "dd-mm-yyyy")
    
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub MyLoad()
Dim cStrSal As String
Dim cStrChq As String


cStrSal = "SELECT T_SAL_RET.code, Sum(T_SAL_RET.TOTAL) AS T_TOTAL, Sum(T_SAL_RET.T_PRICE) AS T_PRICE FROM T_SAL_RET WHERE CODE IS NOT NULL "
If IsDate(xDate1.Text) Then cStrSal = cStrSal & " AND DATE >= " & DateSq(xDate1.Text)
If IsDate(XDATE2.Text) Then cStrSal = cStrSal & " AND DATE >= " & DateSq(xDate1.Text)
cStrSal = cStrSal & " GROUP BY T_SAL_RET.code "
If SalesTable.State = adStateOpen Then SalesTable.Close
SalesTable.Open cStrSal, CON, adOpenKeyset, adLockOptimistic, adCmdText


With grid1
'                           0                           1
    CSTRALL = "  select FILE3_10.CODE AS CODE , FILE3_10.DESCA AS DESCA , " & _
                cField2 & " , " & cField3 & " , " & cField4 & " ,   ' 'AS N5 , ' ' AS N6 , ' '  AS N7 , ' ' AS N8 , " & _
                cField9 & " , " & cField10 & " , " & cField11 & " , " & cField12 & " , ' ' AS N13 " & _
                " FROM (FILE3_11 LEFT JOIN FILE3_10 ON FILE3_11.CODE = FILE3_10.CODE) LEFT JOIN file3_50 ON FILE3_10.GROUP = file3_50.CODE WHERE FILE3_10.CODE IS NOT NULL "
    If xGrCust.BoundText <> "" Then CSTRALL = CSTRALL & " AND [file3_10.GROUP]  = " & MyParn(xGrCust.BoundText)
    If xCode.Text <> "" Then CSTRALL = CSTRALL & " AND FILE3_10.[CODE] = " & MyParn(xCode.Text)
    CSTRALL = CSTRALL & " GROUP BY FILE3_10.DESCA , FILE3_10.CODE "
    DATA2.RecordSource = CSTRALL
    DATA2.Refresh
    
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 14
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    .TextMatrix(0, 2) = "Ã. „»Ì⁄«  »”⁄— «·Ã„·…"
    .TextMatrix(0, 1) = "ﬁÌ„… „»Ì⁄«  ›⁄·Ì…"
    .TextMatrix(0, 1) = "≈Ã„«·Ï «·Œ’„"
    .TextMatrix(0, 1) = "‰”»… «·Œ’„"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xDate1.Text
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
    
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 900
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 800
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 900
    .ColWidth(12) = 1000
    .ColWidth(13) = 900
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
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    End With
End Sub
Sub fillgrd()
    With grid1
    For I = 1 To .Rows - 1
        SalesTable.Find " CODE = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
        If Not SalesTable.EOF Then
            .TextMatrix(I, 5) = Val(.TextMatrix(I, 3)) - Val(.TextMatrix(I, 4))
            .TextMatrix(I, 6) = Format(SalesTable!T_PRICE, "#0.00")
            .TextMatrix(I, 7) = Format(Val(.TextMatrix(I, 6)) - Val(.TextMatrix(I, 5)), "FIXED")
            If Val(.TextMatrix(I, 6)) <> 0 Then .TextMatrix(I, 8) = Format(Val(.TextMatrix(I, 7)) / Val(.TextMatrix(I, 6)) * 100, "FIXED")
            ChqTable.Find " CODE1 = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
            If Not ChqTable.EOF Then
                .TextMatrix(I, 13) = Format(ChqTable!T_VALUE, "FIXED")
            End If
        End If
    Next I
    .SubtotalPosition = flexSTAbove
    For I = 2 To 13
        .Subtotal flexSTSum, -1, I, "#0.00", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 1 Then
        If Val(.TextMatrix(1, 6)) <> 0 Then .TextMatrix(1, 8) = Format(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 6)) * 100, "FIXED")
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub

Private Sub grid1_DblClick()
    ShowSalItemCust.Show 1
End Sub
Private Sub xDOC_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "SELECT FILE7_60H.DOC_NO, file4_10.DESCA, FILE7_60H.date FROM file4_10 RIGHT JOIN FILE7_60H ON file4_10.CODE = FILE7_60H.code "
    Generalarray(2) = "Order by FILE7_60H.DATE DESC "
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·„Ê—œ"
    listarray(0, 1) = "(%%DESCA%%)"
    
    GrdArray(0, 0) = "—ﬁ„ „” ‰œ"
    GrdArray(0, 1) = 1500
    
    GrdArray(1, 0) = "«·«”„"
    GrdArray(1, 1) = 3000
    
    GrdArray(2, 0) = " «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub
Sub myProc()
    xDoc_No.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xDate1.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 2)
    xCodeDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    Unload Search3
Exit Sub
myerror:
Unload Search
End Sub

