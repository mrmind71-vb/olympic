VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form ViewITEMBAL 
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
End
Attribute VB_Name = "ViewITEMBAL"
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
    cHead1 = "√—’œ… √’‰«› ··„Ê—œ " & VsBalItem.grid1.TextMatrix(VsBalItem.grid1.Row, 1)
    Load PrintGrd
    PrintGrd.doprint ItemInv, 1, -1, cHead1, , , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
Dim SalTable As New ADODB.Recordset
cStr1 = " SELECT Sum(FILE1_11.[IN]) AS TIN, Sum(FILE1_11.OUT) AS TOUT,Sum(VAL(FILE1_11.[IN] & '') * VAL(FILE1_10.cost & '')) AS VIN,  Sum(FILE1_11.OUT * FILE1_10.COST) AS VOUT ,Sum(FILE1_11.[IN] * FILE1_10.PRICE) AS VIN2, Sum(FILE1_11.OUT * FILE1_10.price) AS VOUT2 , FILE1_11.ITEM, FILE1_10.DESCA , FILE1_10.cost, FILE1_10.PRICE " & _
        " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.SUPLER = " & MyParn(VsBalItem.grid1.TextMatrix(VsBalItem.grid1.Row, 0))
cStr1 = cStr1 & " GROUP BY FILE1_11.ITEM, FILE1_10.DESCA , FILE1_10.PRICE , FILE1_10.COST "

SalTable.Open cStr1, CON, adOpenStatic, adLockReadOnly, adCmdText
Me.Caption = cHead
With ItemInv
    ItemInv.Cols = 8
    ItemInv.Rows = 1
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 3) = "”⁄—  ﬂ·›…"
    .TextMatrix(0, 4) = "”⁄— „” Â·ﬂ"
    .TextMatrix(0, 5) = " ﬁÌ„ »”⁄— «· ﬂ·›…"
    .TextMatrix(0, 6) = " ﬁÌ„ »”⁄·— «·»Ì⁄"
    .TextMatrix(0, 7) = "‰”»… „” Â·ﬂ"
    
    .RowHeight(0) = 800
    .WordWrap = True
    .ColWidth(0) = 800
    .ColWidth(1) = 2500
    .ColWidth(2) = 800
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ExplorerBar = flexExSortShow

    Do Until SalTable.EOF
        If Val(SalTable!VIN & "") - Val(SalTable!VOUT & "") <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = SalTable!Item
            .TextMatrix(.Rows - 1, 1) = SalTable!Desca & ""
            .TextMatrix(.Rows - 1, 2) = Format(Val(SalTable!TIN & "") - Val(SalTable!TOUT & ""), "Fixed")
            .TextMatrix(.Rows - 1, 3) = Format(Val(SalTable!cost & ""), "Fixed")
            .TextMatrix(.Rows - 1, 4) = Format(Val(SalTable!price & ""), "fixed")
            .TextMatrix(.Rows - 1, 5) = Format(Val(SalTable!VIN & "") - Val(SalTable!VOUT & ""), "Fixed")
            .TextMatrix(.Rows - 1, 6) = Format(Val(SalTable!vIN2 & "") - Val(SalTable!vOUT2 & ""), "Fixed")
             If Val(.TextMatrix(.Rows - 1, 5)) > 0 Then .TextMatrix(.Rows - 1, 7) = Format((Val(.TextMatrix(.Rows - 1, 6)) - Val(.TextMatrix(.Rows - 1, 5))) / Val(.TextMatrix(.Rows - 1, 5)) * 100, "Fixed")
            i = i + 1
        End If
    Loop
    .Subtotal flexSTSum, -1, 2, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 3, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 4, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 5, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "#0", , vbRed, True, " ≈Ã„«·Ï"
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
End With
SalTable.Close
Set SalTable = Nothing
End Sub

