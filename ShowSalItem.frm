VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ShowSalItem 
   Caption         =   " ›’Ì·Ï"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   13020
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3690
      Top             =   4365
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
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   12885
      _cx             =   22728
      _cy             =   7250
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
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
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4095
      Width           =   2580
      Begin VB.CommandButton cmdPrint 
         Height          =   600
         Left            =   1305
         Picture         =   "ShowSalItem.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit 
         Height          =   600
         Left            =   45
         Picture         =   "ShowSalItem.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1230
      End
   End
End
Attribute VB_Name = "ShowSalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aWhere As Variant
Dim cHead As String
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = " ›’Ì·Ï „»Ì⁄«  Ê „— Ã⁄«  ’‰› " & retFlag(aWhere, "ITEM") & "  " & retFlag(aWhere, "DESCA")
    If IsDate(retFlag(aWhere, "DATE1")) Then cHead2 = "„‰ " & retFlag(aWhere, "DATE1")
    If IsDate(retFlag(aWhere, "DATE2")) Then cHead2 = cHead2 & turn(cHead2, " ") & "Õ Ï " & retFlag(aWhere, "DATE2")
    PrintGrd.doprint grid1, 1, -1, cHead1, cHead2, , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
Dim cString As String
'               1                       2               3               4             5                       6     7                                  8                               8                       9              10
cString = " SELECT FILE6_20H.date, FILE3_10.DESCA, FILE6_20H.DOC_NO, SUM(FILE6_20.QUANT)," & _
          " FILE6_20.PRICE , SUM(FILE6_20.DISCOUNT), SUM(FILE6_20.TOTAL) " & _
          " FROM FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE"
If retFlag(aWhere, "ITEM") <> "" Then
    cString = cString & turn(cString) & "FILE6_20.ITEM = " & MyParn(retFlag(aWhere, "ITEM"))
End If

If IsDate(retFlag(aWhere, "DATE1")) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE >= " & addDate(retFlag(aWhere, "ITEM"))
End If

If IsDate(retFlag(aWhere, "DATE2")) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE <= " & addDate(retFlag(aWhere, "ITEM"))
End If
cString = cString & " GROUP BY FILE6_20H.date, FILE3_10.DESCA, FILE6_20H.DOC_NO,FILE6_20.PRICE"
Set grid1.DataSource = data1
data1.ConnectionString = strCon
data1.RecordSource = cString
data1.Refresh
FixGrid
End Sub
Sub FixGrid()
With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = " «—ÌŒ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    .TextMatrix(0, 2) = "—ﬁ„ „” ‰œ"
    .TextMatrix(0, 3) = "„»Ì⁄« "
    .TextMatrix(0, 4) = "”⁄—"
    .TextMatrix(0, 5) = "Œ’„"
    .TextMatrix(0, 6) = "≈Ã„«·Ï ﬁÌ„…"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .ColWidth(6) = 900
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 3, "#0", "#0", vbRed, vbYellow, " "
    .Subtotal flexSTSum, -1, 6, "#0", "#0", vbRed, vbYellow, " "

    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ExplorerBar = flexExSortShow
End With
End Sub




