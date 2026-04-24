VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grdsup1 
   Caption         =   "»Ì«‰«  «·„Â‰œ”Ì‰"
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15045
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   270
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grdsup1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "grdsup1.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdsup1.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "grdsup1.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   10485
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   4470
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
         TabIndex        =   2
         Tag             =   "D"
         Top             =   540
         Width           =   1545
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   540
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   330
         Left            =   135
         TabIndex        =   0
         Top             =   180
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·‘⁄»…"
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   10035
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8880
      Left            =   135
      TabIndex        =   7
      Top             =   1035
      Width           =   14865
      _cx             =   26220
      _cy             =   15663
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
Attribute VB_Name = "grdsup1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
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
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String
Dim aHeader As Variant
cHeader1 = "»Ì«‰ ≈Ã„«·Ì Õ—ﬂ… «·„Ê—œÌ‰ Œ·«· › —… "
If IsDate(xDate1.Text) Or IsDate(xDate2.Text) Then aHeader = AddFlag(aHeader, BetweenString(xDate1.Text, xDate2.Text))
If xGroup.MatchedWithList Then aHeader = AddFlag(aHeader, "„Ã„Ê⁄… „Ê—œÌ‰ " & xGroup.Text)

Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 2)
PrintGrdNew.doprint grid1, 0.9, -2, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 9, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
    openCon con
    DATA1.ConnectionString = strCon
    DATA1.RecordSource = "SELECT * FROM FILE4_50 "
    
    Set xGroup.RowSource = DATA1
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    Fixgrd
    grid1.Rows = 1
End Sub
Private Sub MyLoad()
Dim cString As String
If IsDate(xDate1.Text) Then
    cWhere = " date < " & DateSq(xDate1.Text)
    cField = myiif(cWhere, "[SAL] - [PAY]")
End If
cWhere = ""

If IsDate(xDate1.Text) Then cWhere = " date >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xDate2.Text)

cField = cField & turn(cField, " + ") & _
    myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '1')", "[SAL]") & " AS FIRST_BAL"

cField = cField & "," & _
    myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4')", "[SAL]") & " AS PURCAHSE"

cField = cField & "," & _
        myiif(cWhere & turn(cWhere, " And ") & "(TYPE = '5')", "[PAY]") & " AS R_PURCAHSAE"

cField = cField & "," & _
        myiif(cWhere & turn(cWhere, " And ") & "(TYPE = '4' OR TYPE = 5)", "[SAL]-PAY") & " AS R_PURCAHSAE"

cField = cField & "," & _
         myiif("([TYPE] =  '7' OR [TYPE] =  '8' )", "[PAY]-SAL") & " AS CASH"

cField = cField & "," & _
         myiif("([TYPE] =  'A' OR [TYPE] =  'C')", "SAL - [PAY]") & " AS CHECKS"

cField = cField & "," & _
         myiif(cWhere, "[sal]- [pay] ") & " AS LAST_BAL"

With grid1 '                           0                           1
    cString = "  select FILE4_10.CODE, FILE4_10.DESCA, " & _
                 cField & _
                " FROM FILE4_11 INNER JOIN FILE4_10 ON FILE4_11.CODE = FILE4_10.CODE LEFT JOIN file4_50 ON FILE4_10.[group] = file4_50.CODE"
    If xGroup.BoundText <> "" Then cString = cString & turn(cString) & " file4_10.[group]  = " & MyParn(xGroup.BoundText)
    If xCode.Text <> "" Then cString = cString & turn(cString) & " FILE4_10.[CODE] = " & MyParn(xCode.Text)
    If IsDate(xDate2.Text) Then
        cString = cString & turn(cString) & "FILE4_11.DATE <= " & DateSq(xDate2.Text)
    End If
    
    cString = cString & " GROUP BY FILE4_10.DESCA , FILE4_10.CODE "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
    With grid1

    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xDate1.Text
    .TextMatrix(0, 3) = "Ã. „‘ —Ì« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „‘ —Ì« "
    .TextMatrix(0, 6) = "”œ«œ ‰ﬁœÌ"
    .TextMatrix(0, 7) = "”œ«œ ‘Ìﬂ« "
    .TextMatrix(0, 8) = "«·—’Ìœ " & xDate2.Text
        
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1100
    .ColWidth(7) = 1100
    .ColWidth(8) = 1200
    
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
    StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
    If .Rows > 1 Then
        .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
        .TextMatrix(1, 1) = "«·≈Ã„«·Ì"
        .MergeRow(1) = True
    End If
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub grid1_dblClick()
    If grid1.Row > 1 Then
        supMovefrm.xCode.Text = grid1.TextMatrix(grid1.Row, 0)
        supMovefrm.xDesca.Caption = grid1.TextMatrix(grid1.Row, 1)
        supMovefrm.xDate1.Text = xDate1.Text
        supMovefrm.xDate2.Text = xDate2.Text
        supMovefrm.fillgrd
        supMovefrm.Show
    End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
xCustName.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.Text)) & ""
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
