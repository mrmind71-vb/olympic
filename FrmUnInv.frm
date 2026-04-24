VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form FrmUnInv 
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   4845
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   8115
      _cx             =   14314
      _cy             =   8546
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      FixedCols       =   1
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
Attribute VB_Name = "FrmUnInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataTable As Recordset
Private Sub Form_Load()
cStr1 = " SELECT FILE6_00.DOC_NO, FILE6_00.DATE, FILE6_00.time, FILE6_20.ITEM, FILE6_20.QUANT,  FILE1_10.DESCA " & _
        " FROM (FILE6_00 LEFT JOIN FILE6_20 ON FILE6_00.DOC_NO = FILE6_20.DOC_NO) LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM " & _
        " GROUP BY FILE6_00.DOC_NO, FILE6_00.DATE, FILE6_00.time, FILE6_20.ITEM, FILE6_20.QUANT,  FILE1_10.DESCA "
Set DataTable = mydb.OpenRecordset(cStr1)
With VsItem
    .ExplorerBar = flexExSortShow
    .MergeCells = flexMergeFree
    .Cols = 6
    .Rows = 1
    .TextMatrix(0, 0) = "ÝÇÊæÑÉ"
    .TextMatrix(0, 1) = "ÊÇÑíÎ"
    .TextMatrix(0, 2) = "æÞÊ"
    .TextMatrix(0, 3) = "ßæÏ"
    .TextMatrix(0, 4) = "ÕäÝ"
    .TextMatrix(0, 5) = "ßãíÉ"
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColDataType(1) = flexDTDate
    .ColDataType(2) = flexDTDate
    .ColDataType(5) = flexDTDouble
    If DataTable.RecordCount > 0 Then
        DataTable.MoveFirst
        Do While Not DataTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = DataTable.DOC_NO
            .TextMatrix(.Rows - 1, 1) = Format(DataTable.Date, "DD-MM-YYYY")
            .TextMatrix(.Rows - 1, 2) = Format(DataTable.Time, "SHORT TIME")
            .TextMatrix(.Rows - 1, 3) = TurnValue(DataTable.Item, Null, "")
            .TextMatrix(.Rows - 1, 4) = TurnValue(DataTable.DESCA, Null, "")
            .TextMatrix(.Rows - 1, 5) = TurnValue(DataTable.Quant, Null, "")
            DataTable.MoveNext
        Loop
    End If
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, 0, 5, "##0", , vbRed, RUE, "  "
    .Cell(flexcpAlignment, 0, 5, .Rows - 1, 5) = 7
End With
End Sub
Private Sub VSITEM_DblClick()
    Vs_Inv.xDoc_No.Text = VsItem.TextMatrix(VsItem.Row, 0)
    
    Unload Me
End Sub
