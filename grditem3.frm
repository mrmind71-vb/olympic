VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grditem3 
   Caption         =   "„ «»⁄… „’«—Ì› ”Ì«—…"
   ClientHeight    =   9030
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   15780
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
   ScaleHeight     =   9030
   ScaleWidth      =   15780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6390
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   270
      Width           =   4875
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grditem3.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3555
         Picture         =   "grditem3.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grditem3.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2385
         Picture         =   "grditem3.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   11295
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   4380
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1455
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xCar 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·”Ì«—… :"
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
         Index           =   3
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   8700
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1260
      Top             =   405
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
      Left            =   -1845
      Top             =   -135
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
      Left            =   -1620
      Top             =   -225
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
      Left            =   -1575
      Top             =   -180
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
      Height          =   7395
      Left            =   0
      TabIndex        =   11
      Top             =   1035
      Width           =   15675
      _cx             =   27649
      _cy             =   13044
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   300
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   12
      Top             =   8550
      Visible         =   0   'False
      Width           =   15780
      _ExtentX        =   27834
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grditem3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFilesave As String
Dim con As New ADODB.Connection
Dim oSearchItem As New Search31
Dim LastSalTable As New ADODB.Recordset
Dim LastImpTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmdDelinv_Click()
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1, , , prog1
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
Private Sub CmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï √—’œ… „»Ì⁄«  „‘ —Ì«  ··√’‰«›  "
    If IsDate(xDate1.Text) Then cHead2 = "„‰ : " & Format(xDate1.Text, "YYYY-MM-DD")
    If IsDate(xDate2.Text) Then cHead2 = cHead2 & turn(cHead2, " ") & "Õ Ì : " & Format(xDate2.Text, "YYYY-MM-DD")
    PrintGrd.doprint Me.grid1, 0.8, -4, cHead1, cHead2, , False, False, 8, , Array(1)
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From CARS order by CODE"
    Set xCar.RowSource = data1
    xCar.ListField = "Desca"
    xCar.BoundColumn = "Code"
    
    
    
    Set grid1.DataSource = DATA4
    DATA4.ConnectionString = strCon
    Fixgrd
    grid1.Rows = 1
    LoadText Me
End Sub
Private Sub MyLoad()
Dim cWhere As String
cWhere = ""
If IsDate(xDate1.Text) Then cWhere = " date >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xDate2.Text)

cField = myiif(cWhere & turn(cWhere, " And ") & " ( FLAG = 1)", "[VALUE]") & " AS CHARGES"
cField = cField & "," & _
        myiif(cWhere & turn(cWhere, " And ") & " ( FLAG = 2)", "[VALUE]") & " AS ITEMS"
'                           0               1                 2                3
cString = "  select CAR_MOVE.CAR , CARS.desca," & _
            cField & ",CARS.[PRICE],'' AS RATE " & _
            " from CAR_MOVE INNER JOIN CARS ON CAR_MOVE.CAR = CARS.CODE"

If xCar.MatchedWithList Then cString = cString & turn(cString) & " CAR_MOVE.CAR  = " & xCar.BoundText
With grid1
cString = cString & " GROUP BY CAR_MOVE.CAR, CARS.DESCA,CARS.PRICE"
DATA4.RecordSource = cString
DATA4.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .RowHeight(0) = 700
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·”Ì«—…"
    .TextMatrix(0, 2) = "ﬁÿ⁄ €Ì«—"
    .TextMatrix(0, 3) = "„’«—Ì›"
    .TextMatrix(0, 4) = "”⁄— «·”Ì«—…"
    .TextMatrix(0, 5) = "«·‰”»…"
    .ColFormat(5) = "#.#%"
    .FrozenCols = 2
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    For i = 1 To grid1.Rows - 1
        If Val(grid1.TextMatrix(i, 4)) <> 0 Then
            .TextMatrix(i, 5) = Val(.TextMatrix(i, 3)) / Val(.TextMatrix(i, 4))
        End If
    Next
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
    If .Rows > 1 Then
        .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
        .TextMatrix(1, 1) = "«·≈Ã„«·Ì"
        If Val(grid1.TextMatrix(1, 4)) <> 0 Then
            .TextMatrix(1, 5) = Val(.TextMatrix(1, 3)) / Val(.TextMatrix(1, 4))
        End If
        .MergeRow(1) = True
    End If
    .MergeCells = flexMergeFree

    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Unload Me
Set grditem1 = Nothing
End Sub
Private Sub grid1_dblClick()
'If grid1.Row < 2 Then Exit Sub
'Dim aData As Variant
'aData = AddFlag(Empty, "ITEM", grid1.TextMatrix(grid1.Row, 0))
'aData = AddFlag(aData, "DATE1", xDate1.Text)
'aData = AddFlag(aData, "DATE2", xdate2.Text)
'StoreMove.aData = aData
'StoreMove.Show
End Sub
'Private Sub xDesca_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    FilterGrd grid1, xDesca.Text, 1
'End If
'End Sub
'Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then
'    ItemsLookupAll Me, oSearchItem
'End If
'End Sub
'Private Sub xITEM_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    FilterGrd grid1, xItem.Text, 0
'End If
'End Sub
'Sub myProc()
'xItem.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
'xDesca.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
'Unload oSearchItem
'End Sub
Private Sub xDate1_LostFocus()
myValidDate xDate1
End Sub
Private Sub xdate2_LostFocus()
myValidDate xDate2
End Sub

