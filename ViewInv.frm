VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form ViewInv 
   Caption         =   " ›’Ì·Ï"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   5400
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   10785
      _cx             =   19024
      _cy             =   9525
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      BackColorSel    =   16761024
      ForeColorSel    =   255
      BackColorBkg    =   16777215
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ViewInv.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
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
      WallPaperAlignment=   4
   End
End
Attribute VB_Name = "ViewInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SalTable As Recordset
Dim ItemTable As Recordset
Dim cHead As String
Private Sub Form_Load()
If publicFlag = 1 Then
    If ClientMove.xType.Caption = "4" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file6_20 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption) & " ORDER BY  ITEM  ", dbOpenSnapshot)
        cHead = " „»Ì⁄«  —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    ElseIf ClientMove.xType.Caption = "5" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file6_10 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption) & " ORDER BY ITEM ", dbOpenSnapshot)
        cHead = " „— Ã⁄«   —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    ElseIf ClientMove.xType.Caption = "0" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file8_20 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption), dbOpenSnapshot)
        cHead = "  ”ÊÌ… —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    
    End If
Else
    If ClientMove.xType.Caption = "4" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file7_20 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption) & " ORDER BY  ITEM  ", dbOpenSnapshot)
        cHead = " „‘ —Ì«  —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    ElseIf ClientMove.xType.Caption = "5" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file6_11 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption) & " ORDER BY ITEM ", dbOpenSnapshot)
        cHead = " „— Ã⁄«   —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    ElseIf ClientMove.xType.Caption = "0" Then
        Set SalTable = mydb.OpenRecordset("SELECT * FROM file8_30 WHERE DOC_NO = " & MyParn(ClientMove.xInv.Caption), dbOpenSnapshot)
        cHead = "  ”ÊÌ… —ﬁ„ " & ClientMove.xInv.Caption & " » «—ÌŒ : " & ClientMove.xDateInv.Caption
    
    End If

End If
Set ItemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
Me.Caption = cHead
If ClientMove.xType.Caption = "0" Then
    With ItemInv
        .Cols = 3
        .Rows = 2
        
        .TextMatrix(0, 0) = "„” ‰œ"
        .TextMatrix(0, 1) = "»Ì«‰"
        .TextMatrix(0, 2) = "ﬁÌ„…"
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 6000
        .ColWidth(2) = 1200
        
        SalTable.MoveFirst
        .TextMatrix(1, 0) = TurnValue(SalTable.doc, Null, "")
        .TextMatrix(1, 1) = TurnValue(SalTable.DESCA, Null, "")
        .TextMatrix(1, 2) = Format(SalTable.Value, "#0.00")
        .WordWrap = True
        .RowHeight(1) = 1000

    End With
Else
    With ItemInv
        ItemInv.Cols = 6
        ItemInv.Rows = 1
        .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "⁄»Ê…|" & "ﬂ„Ì…|" & "”⁄—|" & "«·≈Ã„«·Ï"
        .ColWidth(0) = 1200
        .ColWidth(1) = 2500
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        If publicFlag = 2 And Not bopt2 Then .ColHidden(4) = True
        .ColDataType(4) = flexDTDouble
        .ColDataType(5) = flexDTDouble
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ExplorerBar = flexExSortShow
        SalTable.MoveFirst
        i = 1
        Do While True
            .AddItem ""
            .TextMatrix(i, 0) = TurnValue(SalTable.Item, Null, "")
            ItemTable.FindFirst " item = " & MyParn(SalTable.Item)
            If Not ItemTable.NoMatch Then
                .TextMatrix(i, 1) = ItemTable.DESCA
                 .TextMatrix(i, 2) = ItemTable.Pack
                .TextMatrix(i, 3) = QtyToString(SalTable.Quant, ItemTable.Pack)
                .TextMatrix(i, 4) = TurnValue(Format(SalTable.price, "##0.00"), Null, "")
            End If
            .TextMatrix(i, 5) = TurnValue(Format(SalTable.total, "##0.00"), Null, "")
            SalTable.MoveNext
            If SalTable.EOF Then Exit Do
            i = i + 1
        Loop
    End With

End If
End Sub
