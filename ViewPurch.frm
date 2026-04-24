VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form ViewPurch 
   Caption         =   "nm"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_exit 
      Caption         =   "ÎÑæÌ"
      Height          =   510
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4770
      Width           =   2355
   End
   Begin VB.CommandButton CMD_PRINT 
      Caption         =   "ØÈÇÚÉ"
      Height          =   465
      Left            =   8865
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4815
      Width           =   2310
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   4605
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   11040
      _cx             =   19473
      _cy             =   8123
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "ViewPurch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHead As String
Private Sub CMD_EXIT_Click()
Unload Me
End Sub

Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "ÊÝÕíáì ãÔÊÑíÇÊ & ãÑÊÌÚÇÊ  " & grdItem.grid1.TextMatrix(grdItem.grid1.Row, 3)
    cHead2 = BetweenString(Format(grdItem.xdate1.Text, "YYYY-MM-DD"), Format(grdItem.XDATE2.Text, "YYYY-MM-DD"))
    Load PrintGrd
    PrintGrd.doprint ItemInv, 1, -2, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()

cStr1 = " SELECT file1_11.total , file1_11.price , FILE1_11.CODE, FILE1_11.[IN], FILE1_11.OUT, FILE4_10.DESCA, FILE1_11.DOC_ID, FILE1_11.ITEM, FILE1_11.TYPE, FILE1_11.DATE " & _
        " FROM FILE1_11 LEFT JOIN FILE4_10 ON FILE1_11.CODE = FILE4_10.CODE WHERE ( FILE1_11.TYPE = '2' OR  FILE1_11.TYPE = '7' ) AND  FILE1_11.ITEM = " & MyParn(grdItem.grid1.TextMatrix(grdItem.grid1.Row, 2))
If IsDate(grdItem.xdate1.Text) Then cStr1 = cStr1 & " and file1_11.date >= " & DateSq(grdItem.xdate1.Text)
If IsDate(grdItem.XDATE2.Text) Then cStr1 = cStr1 & " and file1_11.date <= " & DateSq(grdItem.XDATE2.Text)
cStr1 = cStr1 & " ORDER BY FILE1_11.DATE "

Dim SalTable As New ADODB.Recordset
SalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
Me.Caption = cHead
With ItemInv
    ItemInv.Cols = 7
    ItemInv.Rows = 1
     
    .FormatString = "ÊÇÑíÎ|" & "ÇáãæÑÏ|" & "ãÓÊäÏ|" & "ãÔÊÑíÇÊ|" & "ãÑÊÌÚ|" & "ÓÚÑ|" & " ÅÌãÇáì"
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColDataType(0) = flexDTDate
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .RowHeight(0) = 400
    .ExplorerBar = flexExSortShow
    Do Until SalTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = Format(SalTable!Date, "YYYY-MM-DD")
        .TextMatrix(.Rows - 1, 1) = TurnValue(SalTable!Desca, Null, "")
        .TextMatrix(.Rows - 1, 2) = TurnValue(SalTable!doc_ID, Null, "")
        .TextMatrix(.Rows - 1, 3) = TurnValue(SalTable!In, Null, "")
        .TextMatrix(.Rows - 1, 4) = TurnValue(SalTable!out, Null, "")
        .TextMatrix(.Rows - 1, 5) = Format(SalTable!price, "#0.00")
        .TextMatrix(.Rows - 1, 6) = Format(SalTable!TOTAL, "#0.00")
        SalTable.MoveNext
    Loop
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 6, "#0.00", , vbRed, True, " ÅÌãÇáì"
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
