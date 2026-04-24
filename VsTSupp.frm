VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form VsTSupp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„ÊœÌ·«  ( «·≈÷«›… - «·„»Ì⁄«  - ‰”»… «·„»Ì⁄«  ) "
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
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
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   150
      Width           =   1785
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
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   585
      Width           =   1785
   End
   Begin VSFlex7LCtl.VSFlexGrid invGrid 
      Height          =   6765
      Left            =   75
      TabIndex        =   4
      Top             =   1650
      Width           =   11715
      _cx             =   20664
      _cy             =   11933
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
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1987
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   1710
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00E3C7AB&
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4050
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1710
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E3C7AB&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   1710
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10890
      TabIndex        =   8
      Top             =   630
      Width           =   735
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10890
      TabIndex        =   7
      Top             =   180
      Width           =   675
   End
   Begin VB.Label xModelDesc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1275
      Width           =   2415
   End
End
Attribute VB_Name = "VsTSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataTable As Recordset
Dim SuppTable As Recordset
Dim TempTable As Recordset
Dim cString As String
Dim PurchTable As Recordset
Dim RetTable As Recordset
Dim SaleTable As Recordset
Dim CostSaleTable As Recordset
Dim BalTable As Recordset

Private Sub Cmd_Print_Click()
Load PrintGrd
PrintGrd.Doprint Me.invGrid, 1, , "≈Ã„«·Ï √—’œ… Ê „»Ì⁄«  «·„’«·‰⁄ ", , , , , 9
PrintGrd.Show 1
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdOk_Click()
Dim nSal As Double
Dim nBal As Double
Dim nRat As Double
Dim cStr1 As String
Dim BalSupp As Recordset
Dim cMosm As String
Me.MousePointer = 11

Me.xBal.Caption = ""

cStr1 = " SELECT FILE1_10.Fact,Max(FILE1_10.OKAZ) AS OKAZ , Sum(FILE7_20.TOTAL) AS T_Total FROM FILE7_20 LEFT JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
If xMosm.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.Text)


cStr1 = cStr1 & " GROUP BY FILE1_10.Fact "
Set PurchTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT FILE1_10.Fact, Sum(FILE6_11.TOTAL) AS T_Total FROM FILE6_11 LEFT JOIN FILE1_10 ON FILE6_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
If xMosm.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.Text)


cStr1 = cStr1 & " GROUP BY FILE1_10.Fact "
Set RetTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT FILE1_10.Fact, Sum(FILE6_20.TOTAL) AS T_Total FROM FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
If xMosm.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.Text)


cStr1 = cStr1 & " GROUP BY FILE1_10.Fact "
Set SaleTable = mydb.OpenRecordset(cStr1)


cStr1 = " SELECT FILE1_10.Fact, Sum(FILE1_10.COST * FILE6_20.QUANT ) AS T_Total FROM FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
If xMosm.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.Text)

cStr1 = cStr1 & " GROUP BY FILE1_10.Fact "
Set CostSaleTable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT FILE1_10.Fact, Sum(FILE1_11![IN] * FILE1_10.COST ) AS T_In ,  Sum(FILE1_11.OUT * FILE1_10.COST ) AS T_Out , Sum(FILE1_11![IN] ) AS Q_In ,  Sum(FILE1_11.OUT ) AS Q_Out FROM FILE1_11 LEFT JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
If xMosm.Text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.Text)

cStr1 = cStr1 & " GROUP BY FILE1_10.Fact "
Set BalTable = mydb.OpenRecordset(cStr1)

With invGrid
.Sort = flexSortNone
Me.MousePointer = 11
xCount.Caption = ""

invGrid.Rows = 1
If PurchTable.RecordCount > 0 Then
    Fillgrd
End If
Me.MousePointer = 0
End With
End Sub
Sub Fillgrd()
Dim nTotal As Recordset

With invGrid
.FixedRows = 1
.FixedCols = 0
.ExplorerBar = flexExSortShow
.Rows = 1
nCount = PurchTable.RecordCount
PurchTable.MoveFirst
.SubtotalPosition = flexSTAbove
i = 0
n1 = 0
n2 = 0
n3 = 0
Do While True
    NPURCH = TurnValue(PurchTable.T_TOTAL, Null, 0)
    RetTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
    If Not RetTable.NoMatch Then NPURCH = NPURCH - TurnValue(RetTable.T_TOTAL, Null, 0)
    If NPURCH <> 0 Then
        
        .AddItem ""
        .RowHeight(.Rows - 1) = 250
        i = i + 1
        .TextMatrix(i, 0) = TurnValue(SayCode(FlagTable, 3, PurchTable.FACT), Null, "")
        .TextMatrix(i, 10) = PurchTable.FACT
        .TextMatrix(i, 1) = Format(PurchTable.T_TOTAL, "#0.00")
        
        RetTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
        If Not RetTable.NoMatch Then .TextMatrix(i, 2) = Format(RetTable.T_TOTAL, "#0.00")
        .TextMatrix(i, 3) = Format(Val(.TextMatrix(i, 1)) - Val(.TextMatrix(i, 2)), "#0.00")
        n1 = n1 + .TextMatrix(i, 3)
        
        CostSaleTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
        If Not CostSaleTable.NoMatch Then .TextMatrix(i, 4) = Format(CostSaleTable.T_TOTAL, "#0.00")
        n3 = n3 + Val(.TextMatrix(i, 4))
        
        SaleTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
        If Not SaleTable.NoMatch Then .TextMatrix(i, 5) = Format(SaleTable.T_TOTAL, "#0.00")
        
        BalTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
        If Not BalTable.NoMatch Then .TextMatrix(i, 6) = Format(TurnValue(BalTable.T_IN, Null, 0) - TurnValue(BalTable.T_OUT, Null, 0), "#0.00")
        n2 = n2 + Val(.TextMatrix(i, 6))
        
        BalTable.FindFirst " fact = " & MyParn(PurchTable.FACT)
        If Not BalTable.NoMatch Then .TextMatrix(i, 7) = Format(TurnValue(BalTable.Q_IN, Null, 0) - TurnValue(BalTable.Q_OUT, Null, 0), "#0")
    '    n2 = n2 + Val(.TextMatrix(i, 6))
        
        If Val(.TextMatrix(i, 4)) > 0 Then .TextMatrix(i, 8) = Format(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(i, 3)) * 100, "#0.00")
            
            
        .TextMatrix(i, 9) = Format(PurchTable.OKAZ, "#0")
    End If
    PurchTable.MoveNext
    If PurchTable.EOF Then Exit Do
Loop

Me.xtPurch.Caption = "«·„‘ —Ì« " & Format(n1, "#0.00")
Me.xTBal.Caption = " ﬁÌ„ «·—’Ìœ" & Format(n2, "#0.00")

.MergeCol(0) = True
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 1, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 2, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 3, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 4, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
.Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"

xCount.Caption = .Rows - 2
.Cell(flexcpAlignment, 1, 0, .Rows - 1, 2) = 7
xRate2.Caption = " ‰”»… «·„»Ì⁄«  " & Format(n3 / n1 * 100, "#0.00")
End With
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub


Private Sub Command2_Click()
    PrnOkazFact.Show 1
End Sub
Private Sub Form_Load()
Set SuppTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
    Data1.DatabaseName = MdbPath
    Data1.RecordSource = "SELECT MOSM, MOSM FROM MOSM  "
    xMosm.ListField = "MOSM"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm


With invGrid
.Cols = 11
.Rows = 1
.RowHeight(0) = 600

.TextMatrix(0, 0) = "«·„’‰⁄"
'.TextMatrix(0, 1) = "ﬁÌ„… „‘ —Ì« "
'.TextMatrix(0, 2) = "ﬁÌ„… „— Ã⁄« "
.TextMatrix(0, 3) = "’«›Ï „‘ —Ì« "
.TextMatrix(0, 4) = "ﬁÌ„…  ﬂ·›… „»Ì⁄« "
.TextMatrix(0, 5) = "ﬁÌ„… „»Ì⁄« "

.TextMatrix(0, 6) = " ﬁÌ„ —’Ìœ „ÊœÌ·« "
.TextMatrix(0, 7) = "⁄œœ ﬁÿ⁄"

.TextMatrix(0, 8) = "‰”»… „»Ì⁄« "
.TextMatrix(0, 9) = "Œ’„ √Êﬂ«“ÌÊ‰"

.ColWidth(0) = 2000
.ColWidth(1) = 0
.ColWidth(2) = 0
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 0
.RowHeight(0) = 500
.WordWrap = True
.ColDataType(1) = flexDTDouble
.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble

.Editable = flexEDNone
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
End With
End Sub
Sub myProc()
    ActiveControl.Text = GrdText(Search.Grid1, 0)
    Unload Search
End Sub
Private Sub xCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SuppTable.FindFirst "fact = " & MyParn(xCode.Text)
    If Not SuppTable.NoMatch Then xCode.Text = SuppTable.CODE
    
    SuppTable.FindFirst "SUPPNAME = " & MyParn(xCode.Text)
    If Not SuppTable.NoMatch Then xCode.Text = SuppTable.CODE
    
    SuppTable.FindFirst "Code = " & MyParn(xCode.Text)
    xCodeDesc.Text = IIf(SuppTable.NoMatch, "", TurnValue(SuppTable.DESCA, Null, ""))
End If
End Sub
Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    Generalarray(3) = "Where DescA Like '%cFilter%'"
    GrdArray(1) = 1000
    GrdArray(2) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub invGrid_DblClick()
    lShowItem = True
    Load VsTItem
    VsTItem.xMosm.Text = Me.xMosm.Text
    VsTItem.xFact.BoundText = invGrid.TextMatrix(invGrid.Row, 0)
    VsTItem.Show 1
End Sub

Private Sub InvGrid_EnterCell()
    With invGrid
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = &HFFFFFF
        .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = &H8000000F
        .Editable = (.Col = 9)
    End With
End Sub
Private Sub invGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim nOkaz As Double
Dim cMosm As String
cMosm = xMosm.Text
With invGrid
        nOkaz = Val(.TextMatrix(.Row, 9))
'        If nOkaz > 0 Then
            cStr1 = " update FILE1_10 SET FILE1_10.OKAZ =  " & nOkaz & " , file1_10.price2 = file1_10.price * " & (100 - nOkaz) / 100 & _
                    " where item is not null "
            cStr1 = cStr1 & " and fact = " & MyParn(.TextMatrix(.Row, 0))
            cStr1 = cStr1 & " and MOSM = " & MyParn(cMosm)
            mydb.Execute cStr1
            
            cStr1 = " update FILE1_20 SET FILE1_20.OKAZ =  " & nOkaz & _
                    " where MODEL is not null "
            cStr1 = cStr1 & " and fact = " & MyParn(.TextMatrix(.Row, 0))
            cStr1 = cStr1 & " and MOSM = " & MyParn(cMosm)
            mydb.Execute cStr1
            MsgBox " „  ⁄œÌ· ‰”»… «·Œ’„ Ê ”⁄— «·√ﬂ«“ÌÊ‰"
'        End If
End With
End Sub
Private Sub Command1_Click()
    Dim DataTable As Recordset
    Dim cMosm As String
    cMosm = xMosm.Text
    For i = 2 To invGrid.Rows - 1
        cStr2 = " SELECT  * FROM FILE1_10 where FACT = " & MyParn(invGrid.TextMatrix(i, 0))
        cStr2 = cStr2 & " and mosm = " & MyParn(cMosm)
        Set DataTable = mydb.OpenRecordset(cStr2)
        With DataTable
        If DataTable.RecordCount > 0 Then
            .MoveFirst
            Do While True And Not .EOF
                nPr = .PRICE2
                nRet = nPr - Fix(nPr)
                nPr = Fix(nPr)
                If nRet > 0.5 Then nPrice2 = nPr + 1
                If nRet < 0.5 And nRet > 0 Then nPrice2 = nPr + 0.5
                If nRet = 0 Then nPrice2 = nPr
                If nRet = 0.5 Then nPrice2 = nPr + 0.5
                .Edit
                .PRICE2 = nPrice2
                .Update
                .MoveNext
            Loop
        End If
        End With
            
    Next i
End Sub

