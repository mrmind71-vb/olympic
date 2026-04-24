VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form TSalItem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ï «·„»Ì⁄« "
   ClientHeight    =   6480
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
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
   ScaleHeight     =   6480
   ScaleWidth      =   9480
   Begin VB.CheckBox xGroup 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ﬁ”Ì„ «·√’‰«› ( „Ã„Ê⁄«  —∆Ì”Ì… - „Ã„Ê⁄«  "
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   3120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   75
      Width           =   915
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
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   160
      Width           =   1440
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
      Height          =   315
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   160
      Width           =   1440
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÿ»«⁄…"
      Height          =   390
      Left            =   2155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   1190
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   105
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   1035
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VSFlex7LCtl.VSFlexGrid InvGrid 
      Height          =   4995
      Left            =   360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8835
      _cx             =   15584
      _cy             =   8811
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   89
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TSalItem.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
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
      Left            =   8640
      TabIndex        =   6
      Top             =   195
      Width           =   675
   End
   Begin VB.Label Label2 
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
      Left            =   6120
      TabIndex        =   5
      Top             =   195
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   870
      Left            =   4350
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   5040
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5250
      Left            =   150
      Top             =   1035
      Width           =   9240
   End
End
Attribute VB_Name = "TSalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset
Dim TempTable As Recordset
Dim cString As String
Dim cString2 As String
Dim cFilter1 As String
Dim cFilter As String
Dim nInd As Byte
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .MGrCode
    TempTable.str2 = .MGrdesc
    TempTable.str3 = .GrCode
    TempTable.str4 = .grdesc
    TempTable.STR5 = .Item
    TempTable.STR6 = .DESCITEM
    
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
'   TempTable.VAL3 = .SUMCOST
    If xDate1.Text = "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  "
    End If
    If xDate1.Text = "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  ≈·Ï " & xDate2.Text
    End If
    If xDate1.Text <> "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text
    End If
    If xDate1.Text <> "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    End If
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\TSalItem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdOk_Click()
'If IsDate(xDate1.Text) And IsDate(xDate2.Text) Then
Me.MousePointer = 11
cFilter = ""
cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS DescItem , Sum(FILE6_20.QUANT) AS SumQUANT, " & _
          " Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum([FILE6_20].[COST]*[FILE6_20].[QUANT]) AS SumCost, " & _
          " First(FILE1_50.DESCA) AS GrDesc, First(FILE1_50.CODE) AS GrCode, First(FILE1_70.CODE) AS MGrCode,  " & _
          " First(FILE1_70.DESCA) AS MGrDesc " & _
          " FROM ((FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON " & _
          " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE" & _
          " WHERE (((FILE6_20.STORE)<>'SS') AND ((FILE1_70.FLAG)=2)) "
If xDate1.Text <> "" Then
    cString = cString & " AND [Date] >= DateValue('" & xDate1.Text & "')"
End If

If xDate2.Text <> "" Then
    cString = cString & " AND [DatE] <= DateValue('" & xDate2.Text & "')"
End If
cString = cString & " GROUP BY FILE1_10.ITEM "
cString = cString & " ORDER BY  First(FILE1_70.CODE), First(FILE1_50.CODE), FILE1_10.ITEM "
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    InvGrid.Rows = 1
End If
Me.MousePointer = 0
End Sub
Sub Fillgrd()
Dim nSubTotal As Double
Dim nTotal As Double
Dim nT As Double
Dim cSubGroup As String, cGroup As String
nSubTotal = 0
nTotal = 0
nT = 0
With InvGrid
.FixedRows = 1
.FixedCols = 1
.ExplorerBar = flexExSortShow
InvGrid.Rows = 1
invtable.MoveFirst
.MergeCells = 2
.SubtotalPosition = flexSTBelow
i = 1
Do
   .AddItem ""
   .TextMatrix(i, 1) = invtable.MGrdesc
   .TextMatrix(i, 2) = invtable.grdesc
   .TextMatrix(i, 3) = invtable.Item
   .TextMatrix(i, 4) = invtable.DESCITEM
   .TextMatrix(i, 5) = Format(invtable.SUMQUANT, "##0.00")
   .TextMatrix(i, 6) = Format(invtable.SUMTOTAL, "##0")
   .MergeCol(1) = True
   .MergeCol(2) = True
   .MergeCol(3) = True
   .MergeCol(4) = True
   .MergeCol(5) = True
   .MergeCol(6) = True
   .MergeRow(i) = True
   invtable.MoveNext
   i = i + 1
Loop Until invtable.EOF
.Sort = 1
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 6, , , RGB(255, 0, 0), True
.Subtotal flexSTSum, 1, 6, , , RGB(155, 100, 0), True
.Subtotal flexSTSum, 2, 6, , , RGB(100, 100, 55), True
.OutlineBar = 1
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlue
cSubGroup = .TextMatrix(1, 2)
cGroup = .TextMatrix(1, 1)

'For I = 1 To .Rows - 4
'    If Not .IsSubtotal(I) Then
'        nSubTotal = nSubTotal + .TextMatrix(I, 6)
'        nTotal = nTotal + .TextMatrix(I, 6)
'        nT = nT + .TextMatrix(I, 6)
'    End If
'    If .TextMatrix(I, 2) <> cSubGroup And .IsSubtotal(I) Then
'        .TextMatrix(I, 6) = Format(nSubTotal, "##0")
'        nSubTotal = 0
'        cSubGroup = .TextMatrix(I, 2)
'    End If
'    If .TextMatrix(I, 1) <> cGroup And .IsSubtotal(I) Then
'        .TextMatrix(I, 6) = Format(nTotal, "##0")
'        nTotal = 0
'        cGroup = .TextMatrix(I, 1)
'    End If
'Next I
'.TextMatrix(.Rows - 1, 6) = Format(nT, "##0")
'.TextMatrix(.Rows - 2, 6) = Format(nTotal, "##0")
'.TextMatrix(.Rows - 3, 6) = Format(nSubTotal, "##0")
End With
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
xDate1.Text = ""
xDate2.Text = ""
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
With InvGrid
.Cols = 7
.Rows = 1
.TextMatrix(0, 1) = "„Ã„Ê⁄… —∆Ì”Ì…"
.TextMatrix(0, 2) = "„Ã„Ê⁄…"
.TextMatrix(0, 3) = "ﬂÊœ"
.TextMatrix(0, 4) = "’‰›"
.TextMatrix(0, 5) = "Ã. „»Ì⁄« "
.TextMatrix(0, 6) = "Ã. ﬁÌ„…"
.ColWidth(0) = 400
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1000
.ColWidth(4) = 2000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
For i = 0 To InvGrid.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.MergeCells = 2
End With
End Sub
Private Sub InvGrid_Click()
If Not InvGrid.IsSubtotal(InvGrid.Row) And (InvGrid.Col = 3 Or InvGrid.Col = 4) Then
    R_Sal_Item (InvGrid.TextMatrix(InvGrid.Row, 3))
End If
If Not InvGrid.IsSubtotal(InvGrid.Row) And InvGrid.Col = 1 Then
    R_Sal_MGr (InvGrid.TextMatrix(InvGrid.Row, 1))
End If

If Not InvGrid.IsSubtotal(InvGrid.Row) And InvGrid.Col = 2 Then
    R_Sal_Gr (InvGrid.TextMatrix(InvGrid.Row, 2))
End If

End Sub
Sub R_Sal_Item(cItem)
cString = " SELECT FILE3_10.DESCA, FILE6_20.DATE, FILE6_20.QUANT, FILE6_20.TOTAL,  " & _
          " FILE6_20.DOC_NO, FILE6_20.PRICE " & _
          " FROM FILE3_10 RIGHT JOIN FILE6_20 ON FILE3_10.CODE = FILE6_20.CODE " & _
          " where file6_20.item = " & MyParn(cItem)
If xDate1.Text <> "" Then
    cString = cString & " AND [Date] >= DateValue('" & xDate1.Text & "')"
End If

If xDate2.Text <> "" Then
    cString = cString & " AND [DatE] <= DateValue('" & xDate2.Text & "')"
End If

cString = cString & " ORDER BY  [date] "
Set invtable = mydb.CreateSnapshot(cString)
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .DOC_NO
    TempTable.str2 = .DESCA
    TempTable.Date1 = !Date
    TempTable.val1 = .Quant
    TempTable.val2 = .price
    TempTable.VAL3 = .TOTAL
    TempTable.str8 = "„»Ì⁄«    " & cItem
    If xDate1.Text = "" And xDate2.Text <> "" Then
        TempTable.str8 = TempTable.str8 & " Õ Ï  " & xDate2.Text
    End If
    If xDate1.Text <> "" And xDate2.Text = "" Then
        TempTable.str8 = TempTable.str8 & " „‰   " & xDate1.Text
    End If
    If xDate1.Text <> "" And xDate2.Text <> "" Then
        TempTable.str8 = TempTable.str8 & " „‰ " & xDate1.Text & " Õ Ï  " & xDate2.Text
    End If
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\SalItem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Sub R_Sal_MGr(cMGr)
cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS DescItem , Sum(FILE6_20.QUANT) AS SumQUANT, " & _
          " Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum([FILE6_20].[COST]*[FILE6_20].[QUANT]) AS SumCost, " & _
          " First(FILE1_50.DESCA) AS GrDesc, First(FILE1_50.CODE) AS GrCode, First(FILE1_70.CODE) AS MGrCode,  " & _
          " First(FILE1_70.DESCA) AS MGrDesc " & _
          " FROM ((FILE1_10 RIGHT JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM) LEFT JOIN FILE1_50 ON " & _
          " FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_70 ON FILE1_50.M_GROUP = FILE1_70.CODE" & _
          " WHERE (((FILE6_20.STORE)<>'SS') AND ((FILE1_70.FLAG)=2)) AND FILE1_70.DESCA = " & MyParn(cMGr)
If xDate1.Text <> "" Then
    cString = cString & " AND [Date] >= DateValue('" & xDate1.Text & "')"
End If

If xDate2.Text <> "" Then
    cString = cString & " AND [DatE] <= DateValue('" & xDate2.Text & "')"
End If
cString = cString & " GROUP BY FILE1_10.ITEM "
cString = cString & " ORDER BY  First(FILE1_70.CODE), First(FILE1_50.CODE), FILE1_10.ITEM "
Set invtable = mydb.CreateSnapshot(cString)
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .MGrCode
    TempTable.str2 = .MGrdesc
    TempTable.str3 = .GrCode
    TempTable.str4 = .grdesc
    TempTable.STR5 = .Item
    TempTable.STR6 = .DESCITEM
    
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
'   TempTable.VAL3 = .SUMCOST
    If xDate1.Text = "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  "
    End If
    If xDate1.Text = "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  ≈·Ï " & xDate2.Text
    End If
    If xDate1.Text <> "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text
    End If
    If xDate1.Text <> "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    End If
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\TSalItem.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Sub R_Sal_Gr(cGr)
cString = "SELECT FILE1_10.ITEM, First(FILE1_10.DESCA) AS ItemDesca, First(FILE1_10.GROUP) AS CodeGr, " & _
          " Sum(FILE6_20.TOTAL) AS SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT " & _
          " FROM (FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) LEFT  JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE " & _
          " WHERE FILE1_50.DESCA = " & MyParn(cGr)
If xDate1.Text <> "" Then
    cString = cString & " AND [Date] >= DateValue('" & xDate1.Text & "')"
End If

If xDate2.Text <> "" Then
    cString = cString & " AND [DatE] <= DateValue('" & xDate2.Text & "')"
End If
cString = cString & " GROUP BY FILE1_10.ITEM "
cString = cString & " ORDER BY   FILE1_10.ITEM "
Set invtable = mydb.CreateSnapshot(cString)
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .CODEGR
    TempTable.str2 = cGr
    TempTable.STR5 = .Item
    TempTable.STR6 = .ItemDescA
    
    TempTable.val1 = .SUMQUANT
    TempTable.val2 = .SUMTOTAL
    If xDate1.Text = "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  "
    End If
    If xDate1.Text = "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  ≈·Ï " & xDate2.Text
    End If
    If xDate1.Text <> "" And xDate2.Text = "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text
    End If
    If xDate1.Text <> "" And xDate2.Text <> "" Then
        TempTable.str8 = "≈Ã„«·Ï „»Ì⁄«  «·√’‰«›  „‰ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
    End If
    TempTable.str19 = firsttitle
    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
Report1.ReportFileName = PublicPath & "\Reports\R_Sal_GR.rpt"
Report1.DataFiles(0) = App.Path & "\temp.mdb"
Report1.Action = 1
End Sub
Private Sub xGroup_Click()
With InvGrid
If xGroup.Value = 1 Then
    .MergeCells = 2
    .Subtotal flexSTClear
    .Subtotal flexSTSum, -1, 6, , , RGB(255, 0, 0), True
    .Subtotal flexSTSum, 1, 6, , , RGB(155, 100, 0), True
    .Subtotal flexSTSum, 2, 6, , , RGB(100, 100, 55), True
    .OutlineBar = 1
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
    .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlue
Else
    .MergeCells = 0
    .Subtotal flexSTClear
'    .Sort
End If
End With
End Sub
