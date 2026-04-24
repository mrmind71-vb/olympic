VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form ViewS_S 
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
Attribute VB_Name = "ViewS_S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SalTable As Recordset
Dim cHead As String
Private Sub CMD_EXIT_Click()
Unload Me
End Sub
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = Me.Caption
    If VsSupp.InvGrid1.Col <> 9 Then cHead2 = "„‰ " & Format(VsSupp.xDate1.Text, "dd-mm-yyyy") & " ≈·Ï  " & Format(VsSupp.xDate2.Text, "dd-mm-yyyy")
    Load PrintGrd
    PrintGrd.doprint ItemInv, 1, -1, cHead1, cHead2, , False, True, 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
Select Case VsSupp.InvGrid1.Col
    Case 3, 4
        Me.Caption = "»Ì«‰ ›Ê« Ì— „‘ —Ì«  - „— Ã⁄«  ··„Ê—œ " & VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 1)
        cStr1 = " SELECT FILE4_11.DOC_ID, FILE4_11.DATE, FILE4_11.PAY, FILE4_11.SAL FROM FILE4_11  " & _
                " WHERE ( FILE4_11.TYPE = '4' OR  FILE4_11.TYPE = '5' ) AND  FILE4_11.CODE = " & MyParn(VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 0))
        If IsDate(VsSupp.xDate1.Text) Then cStr1 = cStr1 & " and file4_11.date >= " & DateSql(VsSupp.xDate1.Text)
        If IsDate(VsSupp.xDate2.Text) Then cStr1 = cStr1 & " and file4_11.date <= " & DateSql(VsSupp.xDate2.Text)
        cStr1 = cStr1 & " ORDER BY FILE4_11.DATE "
        Set SalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
        With ItemInv
            ItemInv.Cols = 4
            ItemInv.Rows = 1
             
            .FormatString = " «—ÌŒ|" & "„” ‰œ|" & "„‘ —Ì« |" & "„— Ã⁄"
            .ColWidth(0) = 1500
            .ColWidth(1) = 1500
            .ColWidth(2) = 2000
            .ColWidth(3) = 2000
            .ColDataType(0) = flexDTDate
            .ColDataType(2) = flexDTDouble
            .ColDataType(3) = flexDTDouble
            .ExplorerBar = flexExSortShow
            If SalTable.RecordCount = 0 Then Exit Sub
            SalTable.MoveFirst
            i = 1
            Do While True
                .AddItem ""
                .TextMatrix(i, 0) = Format(SalTable.Date, "DD-MM-YYYY")
                .TextMatrix(i, 1) = TurnValue(SalTable.Doc_Id, Null, "")
                .TextMatrix(i, 2) = TurnValue(SalTable.sal, Null, 0)
                .TextMatrix(i, 3) = TurnValue(SalTable!PAY, Null, 0)
                SalTable.MoveNext
                If SalTable.EOF Then Exit Do
                i = i + 1
            Loop
            .Subtotal flexSTSum, -1, 2, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
        End With
    Case 6
        Me.Caption = "»Ì«‰ œ›⁄«  ‰ﬁœÏ  ··„Ê—œ " & VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 1)
        cStr1 = " SELECT FILE8_40.DOC_NO, FILE8_40.DATE, FILE8_40.VALUE , FILE8_40.DESCA  FROM FILE8_40  " & _
                " WHERE CODE = " & MyParn(VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 0))
        If IsDate(VsSupp.xDate1.Text) Then cStr1 = cStr1 & " and date >= " & DateSql(VsSupp.xDate1.Text)
        If IsDate(VsSupp.xDate2.Text) Then cStr1 = cStr1 & " and date <= " & DateSql(VsSupp.xDate2.Text)
        cStr1 = cStr1 & " ORDER BY DATE "
        Set SalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
        With ItemInv
            ItemInv.Cols = 4
            ItemInv.Rows = 1
             
            .FormatString = " «—ÌŒ|" & "„” ‰œ|" & "«·»Ì«‰|" & "«·œ›⁄…"
            .ColWidth(0) = 1500
            .ColWidth(1) = 1500
            .ColWidth(2) = 3000
            .ColWidth(3) = 2000
            .ColDataType(0) = flexDTDate
            .ColDataType(3) = flexDTDouble
            .ExplorerBar = flexExSortShow
            If SalTable.RecordCount = 0 Then Exit Sub
            SalTable.MoveFirst
            i = 1
            Do While True
                .AddItem ""
                .TextMatrix(i, 0) = Format(SalTable.Date, "DD-MM-YYYY")
                .TextMatrix(i, 1) = TurnValue(SalTable.doc_no, Null, "")
                .TextMatrix(i, 2) = TurnValue(SalTable.Desca, Null, "")
                .TextMatrix(i, 3) = TurnValue(SalTable!Value, Null, 0)
                SalTable.MoveNext
                If SalTable.EOF Then Exit Do
                i = i + 1
            Loop
            .Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
        End With

    Case 8
        Me.Caption = "»Ì«‰  ”ÊÌ«    ··„Ê—œ " & VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 1)
        cStr1 = " SELECT FILE8_30.DOC_NO, FILE8_30.DATE, FILE8_30.VALUE , FILE8_30.DESCA  FROM FILE8_30  " & _
                " WHERE CODE = " & MyParn(VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 0))
        If IsDate(VsSupp.xDate1.Text) Then cStr1 = cStr1 & " and date >= " & DateSql(VsSupp.xDate1.Text)
        If IsDate(VsSupp.xDate2.Text) Then cStr1 = cStr1 & " and date <= " & DateSql(VsSupp.xDate2.Text)
        cStr1 = cStr1 & " ORDER BY DATE "
        Set SalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
        With ItemInv
            ItemInv.Cols = 4
            ItemInv.Rows = 1
             
            .FormatString = " «—ÌŒ|" & "„” ‰œ|" & "«·»Ì«‰|" & "«·œ›⁄…"
            .ColWidth(0) = 1500
            .ColWidth(1) = 1500
            .ColWidth(2) = 3000
            .ColWidth(3) = 2000
            .ColDataType(0) = flexDTDate
            .ColDataType(3) = flexDTDouble
            .ExplorerBar = flexExSortShow
            If SalTable.RecordCount = 0 Then Exit Sub
            SalTable.MoveFirst
            i = 1
            Do While True
                .AddItem ""
                .TextMatrix(i, 0) = Format(SalTable.Date, "DD-MM-YYYY")
                .TextMatrix(i, 1) = TurnValue(SalTable.doc_no, Null, "")
                .TextMatrix(i, 2) = TurnValue(SalTable.Desca, Null, "")
                .TextMatrix(i, 3) = TurnValue(SalTable!Value, Null, 0)
                SalTable.MoveNext
                If SalTable.EOF Then Exit Do
                i = i + 1
            Loop
            .Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
        End With

    Case 7
        Me.Caption = "»Ì«‰ œ›⁄«  ‘Ìﬂ«     ··„Ê—œ " & VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 1)
        cStr1 = "SELECT FILE5_21.DATE_1, file5_23.DATE, file5_23.VALUE, file5_23.ser_no " & _
                  " FROM FILE5_21 RIGHT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no WHERE FILE5_21.CODE = " & MyParn(VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 0))
        If IsDate(VsSupp.xDate1.Text) Then cStr1 = cStr1 & " and date >= " & DateSql(VsSupp.xDate1.Text)
        If IsDate(VsSupp.xDate2.Text) Then cStr1 = cStr1 & " and date <= " & DateSql(VsSupp.xDate2.Text)
        cStr1 = cStr1 & " ORDER BY DATE "
        Set SalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
        With ItemInv
            ItemInv.Cols = 4
            ItemInv.Rows = 1
             
            .FormatString = " «—ÌŒ|" & "„” ‰œ|" & "«·»Ì«‰|" & "«·œ›⁄…"
            .ColWidth(0) = 1500
            .ColWidth(1) = 1500
            .ColWidth(2) = 4000
            .ColWidth(3) = 2000
            .ColDataType(0) = flexDTDate
            .ColDataType(3) = flexDTDouble
            .ExplorerBar = flexExSortShow
            If SalTable.RecordCount = 0 Then Exit Sub
            SalTable.MoveFirst
            i = 1
            Do While True
                .AddItem ""
                .TextMatrix(i, 0) = Format(SalTable.Date, "DD-MM-YYYY")
                .TextMatrix(i, 1) = TurnValue(SalTable.ser_no, Null, "")
                .TextMatrix(i, 2) = " œ›⁄… „‰ ‘Ìﬂ Õﬁ " & Format(SalTable.date_1, "DD-MM-YYYY")
                .TextMatrix(i, 3) = TurnValue(SalTable!Value, Null, 0)
                SalTable.MoveNext
                If SalTable.EOF Then Exit Do
                i = i + 1
            Loop
            .Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
        End With

    Case 9
        Me.Caption = "‘Ìﬂ«  €Ì— „”œœ…  ··„Ê—œ " & VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 1)
        cStr1 = "SELECT FILE5_21.DATE_1, FILE5_21.VALUE,  FILE5_21.DATE_R, Sum(file5_23.VALUE) AS tpay,max (file5_23.date)  as ldate ,  file5_21.ser_no " & _
                  " FROM FILE5_21 LEFT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no " & _
                  " WHERE file5_21.date_3 is null AND FILE5_21.CODE = " & MyParn(VsSupp.InvGrid1.TextMatrix(VsSupp.InvGrid1.row, 0))
        cStr1 = cStr1 & " GROUP BY FILE5_21.DATE_1, FILE5_21.VALUE, FILE5_21.DATE_R, file5_21.ser_no"
        Set SalTable = mydb.OpenRecordset(cStr1, dbOpenSnapshot)
        
        With ItemInv


            ItemInv.Cols = 6
            ItemInv.Rows = 1
             
            .FormatString = " «—ÌŒ ≈” Õﬁ«ﬁ|" & "„” ‰œ|" & " «—ÌŒ  Õ—Ì—|" & "«·„”œœ"
            .TextMatrix(0, 0) = " «—ÌŒ ≈” Õﬁ«ﬁ"
            .TextMatrix(0, 1) = "„” ‰œ"
            .TextMatrix(0, 2) = " «—ÌŒ  Õ—Ì—"
            .TextMatrix(0, 3) = "ﬁÌ„… «·‘Ìﬂ"
            .TextMatrix(0, 4) = "«·„”œœ"
            .TextMatrix(0, 5) = "«·„ »ﬁÏ"
            
            .ColWidth(0) = 1500
            .ColWidth(1) = 1500
            .ColWidth(2) = 1500
            .ColWidth(3) = 1500
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColDataType(0) = flexDTDate
            .ColDataType(3) = flexDTDouble
            .ColDataType(4) = flexDTDouble
            .ColDataType(5) = flexDTDouble
            .ExplorerBar = flexExSortShow
            If SalTable.RecordCount = 0 Then Exit Sub
            SalTable.MoveFirst
            i = 1
            Do While True
                .AddItem ""
                .TextMatrix(i, 0) = Format(SalTable.date_1, "DD-MM-YYYY")
                .TextMatrix(i, 1) = TurnValue(SalTable.ser_no, Null, "")
                .TextMatrix(i, 2) = Format(SalTable.ldate, "DD-MM-YYYY")
                .TextMatrix(i, 3) = Format(TurnValue(SalTable!Value, Null, 0), "#0.00")
                .TextMatrix(i, 4) = Format(TurnValue(SalTable!tpay, Null, 0), "#0.00")
                .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 3)) - .TextMatrix(i, 4), "#0.00")
                SalTable.MoveNext
                If SalTable.EOF Then Exit Do
                i = i + 1
            Loop
            .Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Subtotal flexSTSum, -1, 4, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Subtotal flexSTSum, -1, 5, "#0.00", , vbRed, True, " ≈Ã„«·Ï"
            .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
        End With
    End Select
End Sub
