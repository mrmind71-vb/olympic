VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ViewITEMSAL 
   Caption         =   " ›’Ì·Ï"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4950
      Width           =   2715
   End
   Begin VB.CommandButton CMD_PRINT 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4950
      Width           =   2715
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   4710
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   11040
      _cx             =   19473
      _cy             =   8308
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
      SelectionMode   =   3
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
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "ViewITEMSAL"
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
    cHead1 = " ›’Ì·Ï Õ—ﬂ… √’‰«› "
    cHead2 = " „‰  «—ÌŒ " & Format(Vstsalsupp.xdate1.Text, "YYYY-MM-DD") & " ≈·Ï  «—ÌŒ " & Format(Vstsalsupp.XDATE2.Text, "YYYY-MM-DD")
    Load PrintGrd
    PrintGrd.doprint ItemInv, 1, -1, cHead1, cHead2, , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
Set xGroupMain.RowSource = DATA2
xGroupMain.ListField = "Desca"
xGroupMain.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
Set xGroup.RowSource = data3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

With ItemInv
    ItemInv.Cols = 8
    ItemInv.Rows = 1
     
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "ﬂ„Ì…"
    .TextMatrix(0, 3) = "„ Ê”ÿ «·»Ì⁄"
    .TextMatrix(0, 4) = "„ Ê”ÿ  ﬂ·›…"
    .TextMatrix(0, 5) = "≈Ã„«·Ï ﬁÌ„… »Ì⁄"
    .TextMatrix(0, 6) = "≈Ã„«·Ï  ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 7) = "‰”»… «·—»Õ"
    
    .ColWidth(0) = 900
    .ColWidth(1) = 2300
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1100
    .ColWidth(7) = 1100

    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ExplorerBar = flexExSortShow
    Do Until SalTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = SalTable!Item
        .TextMatrix(.Rows - 1, 1) = TurnValue(SalTable!Desca, Null, "")
        .TextMatrix(.Rows - 1, 2) = Format(SalTable!TQUANT, "#0")
        If TurnValue(SalTable!TQUANT, Null, 0) <> 0 Then .TextMatrix(.Rows - 1, 3) = Format(SalTable!TTOTAL / SalTable!TQUANT, "#0.00")
        If TurnValue(SalTable!TQUANT, Null, 0) <> 0 Then .TextMatrix(.Rows - 1, 4) = Format(SalTable!Tcost / SalTable!TQUANT, "#0.00")
        .TextMatrix(.Rows - 1, 5) = Format(SalTable!TTOTAL, "#0.00")
        .TextMatrix(.Rows - 1, 6) = Format(SalTable!Tcost, "#0.00")
        If Val(.TextMatrix(.Rows - 1, 6)) > 0 Then .TextMatrix(.Rows - 1, 7) = Format((Val(.TextMatrix(.Rows - 1, 5)) - Val(.TextMatrix(.Rows - 1, 6))) / Val(.TextMatrix(.Rows - 1, 6)) * 100, "#0.00")
        SalTable.MoveNext
        I = I + 1
    Loop
    .Subtotal flexSTSum, -1, 2, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 5, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
End With
End Sub
Private Sub ItemInv_DBLClick()
Dim cItem As String
cItem = ItemInv.TextMatrix(ItemInv.Row, 0)
cdesca = ItemInv.TextMatrix(ItemInv.Row, 1)
cString = " SELECT FILE3_10.DESCA, FILE6_20H.DATE , FILE6_20.QUANT, FILE6_20.QUANT * FILE6_20.PRICE AS TOTAL,   " & _
          " FILE6_20.DOC_NO, FILE6_20.PRICE " & _
          " FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) LEFT JOIN FILE3_10 ON FILE3_10.CODE = FILE6_20H.CODE " & _
          " where file6_20.item = " & MyParn(cItem)

If IsDate(Vstsalsupp.xdate1.Text) Then cString = cString & " AND file6_20H.Date >= " & DateSq(Vstsalsupp.xdate1.Text)
If IsDate(Vstsalsupp.XDATE2.Text) Then cString = cString & " AND file6_20H.Date <= " & DateSq(Vstsalsupp.XDATE2.Text)
cString = cString & " ORDER BY  file6_20.Date "
Dim invTable As New ADODB.Recordset
invTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim temptable As New ADODB.Recordset
contemp.Execute "Delete * From Temp"
temptable.Open "Temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With invTable
Do Until .EOF
    temptable.AddNew
    temptable!str2 = !doc_no
    temptable!str3 = !Desca
    temptable!Date1 = !Date
    temptable!val1 = !Quant
    temptable!val2 = !price
    temptable!val3 = !TOTAL
    
    'temptable!str8 = "„»Ì⁄«    " & citem & " ==> " & ItemInv.TextMatrix(ItemInv.Row, 0)
    temptable!str8 = "„»Ì⁄«    " & cdesca
    temptable!str9 = " „‰  «—ÌŒ  " & Vstsalsupp.xdate1.Text & " Õ Ï  " & Vstsalsupp.XDATE2.Text
    temptable!str19 = Firsttitle
    temptable!STR20 = TurnValue(Secondtitle)
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = PublicPath & "\Reports\SalItem.rpt"
main.Report1.DataFiles(0) = cPathTemp
main.Report1.Action = 1
End Sub




