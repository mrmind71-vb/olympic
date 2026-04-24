VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form ViewSale 
   Caption         =   "ÊÝÕíáì"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_exit 
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4815
      Width           =   1950
   End
   Begin VB.CommandButton CMD_PRINT 
      Caption         =   "ØÈÇÚÉ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4815
      Width           =   1815
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   4620
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   11040
      _cx             =   19473
      _cy             =   8149
      _ConvInfo       =   1
      Appearance      =   1
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
Attribute VB_Name = "ViewSale"
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
    cHead1 = "ÊÝÕíáì ãÈíÚÇÊ & ãÑÊÌÚÇÊ  " & grdItem.grid1.TextMatrix(grdItem.grid1.Row, 3) & cHead2 & BetweenString(grdItem.xdate1.Text, Format(grdItem.XDATE2.Text, "YYYY-MM-DD"))
    Load PrintGrd
    PrintGrd.doprint ItemInv, 1, -2, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
cStr1 = " SELECT FILE3_10.DESCA, FILE1_11.DATE, FILE1_11.[IN], FILE1_11.OUT, FILE1_11.PRICE, FILE1_11.DOC_ID  , file1_11.total  FROM FILE1_11 RIGHT JOIN FILE3_10 ON FILE1_11.CODE = FILE3_10.CODE " & _
        " WHERE (FILE1_11.TYPE = '6' OR  FILE1_11.TYPE = '3' ) AND  FILE1_11.ITEM = " & MyParn(grdItem.grid1.TextMatrix(grditem1.grid1.Row, 2))
If IsDate(grditem1.xdate1.Text) Then cStr1 = cStr1 & " and file1_11.date >= " & DateSq(grditem1.xdate1.Text)
If IsDate(grditem1.XDATE2.Text) Then cStr1 = cStr1 & " and file1_11.date <= " & DateSq(grditem1.XDATE2.Text)
cStr1 = cStr1 & " ORDER BY FILE1_11.DATE"
Dim SalTable As New ADODB.Recordset
SalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

Me.Caption = cHead
With ItemInv
    ItemInv.Cols = 7
    ItemInv.Rows = 1
     
    .FormatString = "ÊÇÑíÎ|" & "ÇáÚãíá|" & "ãÓÊäÏ|" & "ãÈíÚÇÊ|" & "ãÑÊÌÚ|" & "ÓÚÑ|" & "ÅÌãÇáì"
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .RowHeight(0) = 1000
    .ColDataType(0) = flexDTDate
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ExplorerBar = flexExSortShow
    Do Until SalTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = Format(SalTable!Date, "YYYY-MM-DD")
        .TextMatrix(.Rows - 1, 1) = TurnValue(SalTable!Desca, Null, "")
        .TextMatrix(.Rows - 1, 2) = TurnValue(SalTable!doc_ID, Null, "")
        .TextMatrix(.Rows - 1, 3) = TurnValue(SalTable!out, Null, "")
        .TextMatrix(.Rows - 1, 4) = TurnValue(SalTable!In, Null, "")
        .TextMatrix(.Rows - 1, 5) = Format(SalTable!price, "#0.00")
        .TextMatrix(.Rows - 1, 6) = Format(SalTable!TOTAL, "#0.00")
        SalTable.MoveNext
    Loop
    .Subtotal flexSTSum, -1, 6, "#0.00", , vbRed, True, " ÅÌãÇáì"
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
End With
End Sub
