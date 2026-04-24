VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form Dream_Bar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CommandButton Command5 
      Caption         =   "≈÷«›… —’Ìœ „ÊœÌ·«  „Ê—œ"
      Height          =   540
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -75
      Width           =   1875
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… „— Ã⁄"
         Height          =   540
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ÿ»«⁄… «” Ìﬂ—“"
         Height          =   540
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
         Height          =   540
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   6540
      Left            =   75
      TabIndex        =   0
      Top             =   855
      Width           =   11565
      _cx             =   20399
      _cy             =   11536
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7425
      Width           =   3990
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   375
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„Êœ :"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7425
      Width           =   6315
      Begin VB.CommandButton cmduno 
         Caption         =   " —«Ã⁄"
         Height          =   540
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
         Height          =   540
         Left            =   4875
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·„ÿ»Ê⁄"
         Height          =   540
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelAll 
         Caption         =   "Õ–› «·ﬂ·"
         Height          =   540
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Œ—ÊÃ"
         Height          =   540
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ï ⁄œœ «·ﬁÿ⁄ ··ÿ»«⁄…"
      Height          =   390
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   150
      Width           =   2040
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   150
      Width           =   1065
   End
End
Attribute VB_Name = "Dream_Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Dim tAddPrint As New adodb.Recordset
Dim ItemTable As New adodb.Recordset
Dim nCols As Double
Dim NROWS As Double
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
myreplace
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
myLoad
End Sub
Private Sub cmduno_Click()
myLoad
End Sub

Private Sub Command1_Click()
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    delcheck
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
frmReturn.Show 1
myLoad
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    grid1.Rows = 1
    grid1.AddItem ""
    con.Execute "delete  from addprint where addprint.desca = '1'"
End If
End Sub
Private Sub Command4_Click()
If Val(xRow.Text) > 14 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If
If Val(xCol.Text) > 4 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If

myreplace
DoprintStr
Set myForm = Me
CardPrintNew.Show 1
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    delcheck
End If
End Sub
Private Sub Command5_Click()
    BalBar.Show 1
    tAddPrint.Requery
    myLoad
End Sub
Private Sub Form_Load()
'Me.Keyboard1.EnglishKeyboard
openCon con
tAddPrint.Open "Select addPrint.Item,addPrint.DescA,addPrint.Price,addPrint.Quant,addprint.isPrint,AddPrint.Doc_No,AddPrint.SERITEM From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item where addprint.desca = '1' order by DOC_NO , SERITEM ", con, adOpenKeyset, adLockReadOnly
ItemTable.Open "file1_10", con, adOpenKeyset, adLockReadOnly, adCmdTable

With grid1
    .Cols = 7
    .Rows = 2
    .TextMatrix(0, 0) = "—ﬁ„ «·’‰› "
    .TextMatrix(0, 1) = "«·’‰‹‹‹‹‹›"
    .TextMatrix(0, 2) = "«·”⁄—"
    .TextMatrix(0, 3) = "«·ﬂ„Ì… "
    .TextMatrix(0, 4) = "ÿ»«⁄…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1000
    .ColHidden(5) = True
    .ColHidden(6) = True
    
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColDataType(4) = flexDTBoolean
    .ColComboList(0) = "..."
End With
myLoad
grid1.Row = 1
grid1.Col = 0
End Sub
Sub myLoad()
Dim nTotal As Double
With grid1
.Rows = 1
tAddPrint.Requery

Do Until tAddPrint.EOF
   .AddItem ""
    ItemTable.Find "item = " & MyParn(tAddPrint!Item & ""), , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
        .TextMatrix(grid1.Rows - 1, 0) = TurnValue(tAddPrint!Item, Null, "")
        If Val(ItemTable!price & "") = O Then MsgBox " ·« ÌÊÃœ ”⁄— „” Â·ﬂ  " & ItemTable!desca & " ===>  " & ItemTable!Item
        .TextMatrix(grid1.Rows - 1, 1) = ItemTable!desca & ""
        .TextMatrix(grid1.Rows - 1, 3) = tAddPrint!Quant & ""
        .TextMatrix(grid1.Rows - 1, 2) = ItemTable!price & ""
        .TextMatrix(grid1.Rows - 1, 4) = IIf(tAddPrint!isPrint, "-1", "0")
        .TextMatrix(grid1.Rows - 1, 5) = tAddPrint!doc_no & ""
        .TextMatrix(grid1.Rows - 1, 6) = tAddPrint!SERITEM & ""
    End If
    nTotal = nTotal + Val(.TextMatrix(grid1.Rows - 1, 3))
    tAddPrint.MoveNext
Loop
.AddItem ""
End With
Me.xTotal.Caption = nTotal
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tAddPrint.Close
ItemTable.Close
Set tAddPrint = Nothing
Set ItemTable = Nothing
closeCon con
Unload Search3
Unload Me
Set Dream_Bar = Nothing
End Sub
Private Sub GRID1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookupAll Me, Search3
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 3 Or grid1.Col = 4 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub GRID1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem grid1.Row
    End If
End If
If KeyCode = 112 And grid1.Col = 0 Then
    'ItemsLookup
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    'ItemsLookup
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
End If
End Sub
Private Sub myreplace()
con.Execute "Delete  From addprint where desca = '1'"
With grid1
For i = 1 To .Rows - 2
    If Not grid1.RowHidden(i) Then
        If .TextMatrix(i, 0) <> "" Then
            cString = "insert into ADDPRINT( QUANT,ISPRINT,item,[DESCA],DOC_NO,SERITEM)" & _
            " Values(" & _
            Val(.TextMatrix(i, 3)) & "," & _
            IIf(Val(.TextMatrix(i, 4)) = 0, "0", "1") & "," & _
            addstring(.TextMatrix(i, 0)) & "," & _
            addstring("1") & "," & _
            addstring(.TextMatrix(i, 5)) & "," & _
            addvalue(.TextMatrix(i, 6)) & _
            ")"
            con.Execute cString
        End If
    End If
Next
End With
'MyLoad
End Sub
Private Function MYVALID() As Boolean
With grid1
For i = 1 To grid1.Rows - 2
    If Val(.TextMatrix(i, 3)) = 0 Then
        MsgBox "«·ﬂ„Ì… €Ì— „”Ã·…"
        Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Sub delcheck()
For i = 1 To grid1.Rows - 2
   If Val(grid1.TextMatrix(i, 4)) <> 0 Then
        grid1.RowHidden(i) = True
   End If
Next
myreplace
End Sub
Sub myproc()
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 0)
    If nFound <> -1 Then
'        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        MsgBox ("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound)
        Exit Sub
    End If
        
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    grid1.TextMatrix(grid1.Row, 4) = "-1"
    GrdDesc grid1.Row
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = 1
'        Grid1.TextMatrix(Grid1.Rows, 4) = "-1"
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 3) = 1
        grid1.TextMatrix(grid1.Row, 4) = "-1"
        grid1.Select grid1.Rows - 1, 0
    End If
'    CalcTotals
End If
End Sub
Private Sub GrdDesc(nRow)
With grid1
ItemTable.Find "item = " & MyParn(grid1.TextMatrix(nRow, 0)), , adSearchForward, adBookmarkFirst
If ItemTable.EOF Then
    grid1.TextMatrix(nRow, 1) = ""
    grid1.TextMatrix(nRow, 2) = ""
    grid1.TextMatrix(nRow, 3) = ""
    grid1.TextMatrix(nRow, 4) = ""
Else
    .TextMatrix(nRow, 1) = ItemTable!desca & ""
    .TextMatrix(nRow, 2) = ItemTable!price & ""
End If
End With
End Sub
Private Sub DoprintStr()
Dim tCard As New adodb.Recordset
Dim tPrint As New adodb.Recordset
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.4)
SettingArray(cLeftMargin) = MyMeasure(0.5)
SettingArray(cCardWidth) = MyMeasure(5.22)
SettingArray(cCardHeight) = MyMeasure(2.12)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 14
SettingArray(cCols) = 4
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
tPrint.Open "SELECT addPrint.Item,FILE1_10.DESCA,FILE1_10.Price,FILE1_10.Price2,addPrint.Quant  From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint = 1 and addprint.desca = '1' ORDER BY DOC_NO , SERITEM ", con, adOpenKeyset, adLockReadOnly, adCmdText

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * nCols) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡

nadd = MyMeasure(0.2)
Do Until tPrint.EOF
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = False
        
        nHeight = 0
        nLast = MyMeasure(0.1)
        nLastCol = MyMeasure(0.2)
        For nCount = 1 To 1
            tCard.AddNew
            tCard!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.3) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(4.5)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!TextAlign = taCenterTop
            tCard!fontsize = 7
            tCard!Text = TurnValue(Trim(tPrint!desca & ""))
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(1.1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.65) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.6)
            tCard!Height = MyMeasure(0.35) + MyMeasure(0.4)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 8
            tCard!TextAlign = taCenterTop
            tCard!Text = tPrint!Item
            tCard!isBarcode = True
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.1) + MyMeasure(0.4)
            tCard!Width = MyMeasure(4.5)
            tCard!TextAlign = taCenterTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 9
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update
    
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.05) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.1) - MyMeasure(0.37)
            tCard!Width = MyMeasure(0.9)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 10
            If Val(tPrint!price & "") <> 0 Then
                If Val(tPrint!price & "") <> Val(tPrint!PRICE2 & "") Then
                    tCard!Text = myNear(Format(Val(tPrint!price) * 1.5, "#0.00"), 0.5)
                Else
                    tCard!Text = Val(tPrint!price & "")
                End If
            End If
            tCard!TextAlign = taCenterTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.05) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.99) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.37)
            tCard!Width = MyMeasure(0.9)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 10
            If Val(tPrint!price & "") <> 0 Then
                tCard!Text = "_____"
            End If
            tCard!TextAlign = taCenterTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.05) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.1) + nHeight - IIf(nRow <> 1, nLast, 0) - MyMeasure(0.1) + MyMeasure(0.4) - MyMeasure(0.37)
            tCard!Width = MyMeasure(0.9)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!fontsize = 10
            If Val(tPrint!price & "") <> 0 Then
                tCard!Text = (tPrint!price * 12)
            End If
            tCard!TextAlign = taCenterTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            nHeight = SettingArray(cCardHeight) / 1
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop
End With
End Sub
Private Function myvalidRowCol() As Boolean
'If Val(xRow.Text) > SettingArray(cRows) Then
If Val(xRow.Text) > NROWS Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Function
End If

If Val(xCol.Text) > nCols Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Function
End If
myvalidRowCol = True
End Function
Private Sub DoprintStr2(lPrice)
Dim tCard As New adodb.Recordset
Dim tPrint As New adodb.Recordset
Dim cPrice As String
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(1.2)
SettingArray(cLeftMargin) = MyMeasure(1.2)
SettingArray(cCardWidth) = MyMeasure(3.81)
SettingArray(cCardHeight) = MyMeasure(2.1125)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 13
SettingArray(cCols) = 5
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete * From Card"

tCard.Open "card", con, adOpenKeyset, adLockOptimistic, adCmdTable
tPrint.Open "SELECT addPrint.Item,FILE1_10.Price,addPrint.Quant From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint and desca = '1'", con, adOpenKeyset, adLockReadOnly, adCmdText

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * nCols) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
Do Until tPrint.EOF
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastrow = False
'        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            tCard.AddNew
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.32) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.45)
            tCard!Height = MyMeasure(0.6)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!fontsize = 6
            tCard!TextAlign = taRightTop
            tCard!Text = tPrint!Item
            tCard!isBarcode = True
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.5)
            tCard!Height = MyMeasure(0.8)
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterMiddle
            tCard!Text = "ss"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.76) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!fontsize = 8
            tCard!TextAlign = taCenterTop
            'tCard!Text = tPrint!Desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.05) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(3.5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = False
            tCard!FontUnderline = False
            tCard!fontsize = 7
            tCard!TextAlign = taCenterTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


            'tCard.AddNew
            'tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            'tCard!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            'tCard!Width = MyMeasure(2)
            'tCard!TextAlign = taRightTop
            'tCard!Height = 0
            'tCard!FontName = "arial"
            'tCard!FontBold = True
            'tCard!FontSize = 8
            'tCard!Text = tPrint!Item
            'tCard!CardNo = nCard
            'tCard.Update
                        
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop
End With
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub foundOther()
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Sub
    End If
Next
End Sub

