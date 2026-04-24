VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form Morsh_Bar 
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
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   5115
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -75
      Width           =   6525
      Begin VB.CommandButton CMD_Bar56 
         Caption         =   "ÿ»«⁄… 56"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.Label xTotal56 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   345
         Width           =   765
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
      TabIndex        =   9
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   5325
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7425
      Width           =   6315
      Begin VB.CommandButton cmduno 
         Caption         =   " —«Ã⁄"
         Height          =   540
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
         Height          =   540
         Left            =   4875
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "≈Ã„«·Ï ⁄œœ «·ﬁÿ⁄ ··ÿ»«⁄…"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   225
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "Morsh_Bar"
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
MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
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
Private Sub Command3_Click()
frmReturn.Show 1
myLoad
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    grid1.Rows = 1
    grid1.AddItem ""
    con.Execute "delete from addprint"
End If
End Sub
Private Sub CMD_Bar56_Click()
    If Val(xRow.Text) > 14 Then
        MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
        Exit Sub
    End If
    
    If Val(xCol.Text) > 4 Then
        MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
        Exit Sub
    End If
    
    myreplace
    DoprintStr56
    Set myForm = Me
    CardPrintNew.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command5_Click()
    BalBar.Show 1
    tAddPrint.Requery
    myLoad
End Sub

Private Sub Command6_Click()
    If Val(xRow.Text) > 24 Then
        MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
        Exit Sub
    End If
    
    If Val(xCol.Text) > 6 Then
        MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
        Exit Sub
    End If
    
    myreplace
    DoprintStr144
    Set myForm = Me
    CardPrintNew128.Show 1
'    If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
'        delcheck
'    End If

End Sub
Private Sub Form_Load()
openCon con
tAddPrint.Open "Select addPrint.Item,file1_10.DescA , file1_10.Price, file1_10.Price2,addPrint.Quant,addprint.isPrint,addprint.isPriCE,AddPrint.Doc_No ,AddPrint.code , file1_10.t_bar From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item order by file1_10.ITEM ", con, adOpenKeyset, adLockReadOnly
ItemTable.Open "file1_10", con, adOpenKeyset, adLockReadOnly, adCmdTable
With grid1
    .Cols = 8
    .Rows = 2
    .TextMatrix(0, 0) = "—ﬁ„ «·’‰› "
    .TextMatrix(0, 1) = "«·’‰‹‹‹‹‹›"
    .TextMatrix(0, 2) = "«·”⁄—"
    .TextMatrix(0, 3) = "«·ﬂ„Ì… "
    .TextMatrix(0, 4) = "ÿ»«⁄…"
    .TextMatrix(0, 5) = "‰Ê⁄ "
    .TextMatrix(0, 7) = "ÿ»«⁄… ”⁄— «·Ã„·…"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1000
    .ColWidth(5) = 600
    .ColWidth(7) = 800
    .ColHidden(6) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColDataType(4) = flexDTBoolean
    .ColDataType(7) = flexDTBoolean
End With
myLoad
grid1.Row = 1
grid1.Col = 0
End Sub
Sub myLoad()
Dim nTotal144 As Double
Dim nTotal56 As Double

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
        .TextMatrix(grid1.Rows - 1, 7) = IIf(tAddPrint!isPRICE, "-1", "0")
        .TextMatrix(grid1.Rows - 1, 5) = tAddPrint!T_BAR & ""
        .TextMatrix(grid1.Rows - 1, 6) = tAddPrint!doc_no & ""
        If .TextMatrix(grid1.Rows - 1, 5) = "1" Then
            nTotal56 = nTotal56 + Val(.TextMatrix(grid1.Rows - 1, 3))
        Else
            nTotal144 = nTotal144 + Val(.TextMatrix(grid1.Rows - 1, 3))
        End If
    End If
    tAddPrint.MoveNext
Loop
.AddItem ""
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
tAddPrint.Close
ItemTable.Close
Set tAddPrint = Nothing
Set ItemTable = Nothing
Unload Search3
Unload Me
Set Dream_Bar = Nothing
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 5 Or grid1.Col = 7 Then
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
    ItemsLookupAll Me, Search3
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
con.Execute "Delete From addprint"
With grid1
For i = 1 To .Rows - 2
    If Not grid1.RowHidden(i) Then
        If .TextMatrix(i, 0) <> "" Then
            cString = "insert into ADDPRINT( QUANT,TYPE_BARCODE,DOC_NO,ISPRINT,ISPRICE,item)" & _
            " Values(" & _
            Val(.TextMatrix(i, 3)) & "," & _
            Val(.TextMatrix(i, 5)) & "," & _
            addstring(.TextMatrix(i, 6)) & "," & _
            Val(.TextMatrix(i, 4)) & "," & _
            Val(.TextMatrix(i, 7)) & "," & _
            addstring(.TextMatrix(i, 0)) & _
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
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
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
Private Sub DoprintStr56()
Dim tCard As New adodb.Recordset
Dim tPrint As New adodb.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(0.25)
SettingArray(cCardWidth) = MyMeasure(5.25)
SettingArray(cCardHeight) = MyMeasure(2.1)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 14
SettingArray(cCols) = 4
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText
'tPrint.Open "Select addPrint.isPRICE , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE ,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant , addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint = 1 AND addPrint.TYPE_BARCODE = 1 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText
tPrint.Open "Select addPrint.isPRICE , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE ,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant , addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint = 1 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText

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
If tPrint.RecordCount = 0 Then Exit Sub
Do
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = nCols)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            
            '≈”„ «·„Õ·
            tCard.AddNew
            tCard!Top = MyMeasure(0.25) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(1.7) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Shurooq 07"
            tCard!FontBold = 1
            tCard!FontUnderline = 0
            tCard!fontsize = 16
            tCard!TextAlign = taCenterTop
            tCard!Text = "„ﬂ »… √„Ì‰"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            ' ·Ì›Ê‰
            tCard.AddNew
            tCard!Top = MyMeasure(0.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!fontsize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = "Tel." & cHPhone1
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Top = MyMeasure(0.55) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!fontsize = 7
            tCard!TextAlign = taLeftTop
'            tCard!Text = cHPhone2
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.3)
            tCard!Height = MyMeasure(0.45)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!isBarcode = 1
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            

' ITEM
            tCard.AddNew
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(3.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!fontsize = 8
            tCard!TextAlign = taRightTop
            If Len(tPrint!Item) < 10 Then
                tCard!Text = tPrint!Item
            Else
'                tCard!Text = "/" & DelZero(tPrint!code) & "/"
            End If
            tCard!FontUnderline = 1

            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            
            
' DESCA
            tCard.AddNew
            tCard!Top = MyMeasure(1.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(4.8)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!fontsize = 9
            tCard!TextAlign = taCenterBottom
            tCard!Text = tPrint!desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


' PRICE
            tCard.AddNew
            tCard!Top = MyMeasure(1.7) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!fontsize = 10
            tCard!TextAlign = taLeftTop
            If tPrint!isPRICE Then
                tCard!Text = "L.E. " & Format(tPrint!PRICE2, "Fixed")
                tCard!FontUnderline = 1
            Else
                tCard!Text = "L.E. " & Format(tPrint!price, "Fixed")
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' COST
            If Val(tPrint!cost & "") > 0 Then
                tCard.AddNew
                tCard!Top = MyMeasure(1.7) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
                tCard!Left = MyMeasure(3.6) - IIf(blastcol, nLastCol, 0)
                tCard!Width = MyMeasure(1)
                tCard!Height = MyMeasure(0)
                tCard!FontName = "arial"
                tCard!FontBold = 0
                tCard!fontsize = 8
                tCard!TextAlign = taRightTop
                
                nCost = Val(LastPrice(tPrint!Item) & "")
                If nCost = 0 Then nCost = tPrint!cost
                If tPrint!isPRICE Then
                    nCost = (Val(tPrint!PRICE2 & "") - Val(tPrint!cost & "")) * 100 / 2
                Else
                    If Val(tPrint!package & "") <> 0 Then nCost = (Val(tPrint!price & "") - (nCost / Val(tPrint!package & ""))) / 2 * 100
                End If
                If IsNull(tPrint!doc_no) Then
                    tCard!Text = Format(nCost, "#0")
                Else
                    tCard!Text = Format(nCost, "#0") & "/" & DelZero(tPrint!doc_no)
                End If
                tCard!ForeColor = vbBlack
                tCard!CardNo = nCard
                tCard.Update
            End If
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Sub DoprintStr144()
Dim tCard As New adodb.Recordset
Dim tPrint As New adodb.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(1.33)
SettingArray(cLeftMargin) = MyMeasure(0.3)
SettingArray(cCardWidth) = MyMeasure(3.48)
SettingArray(cCardHeight) = MyMeasure(1.23)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 23
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete  From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText
'tPrint.Open "Select addPrint.isPRICE , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant,addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  = 1 AND addPrint.TYPE_BARCODE = 2 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText
tPrint.Open "Select addPrint.isPRICE , file1_10.ITEM,file1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2 ,FILE1_10.COST , FILE1_10.Package ,addPrint.Quant,addPrint.doc_no From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  = 1 ORDER BY FILE1_10.ITEM ", con, adOpenKeyset, adLockOptimistic, adCmdText

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
If tPrint.RecordCount = 0 Then Exit Sub
Do
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = nCols)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0)
        For nCount = 1 To 1
            
            '≈”„ «·„Õ·
            tCard.AddNew
            tCard!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(1.5) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1.8)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Shurooq 07"
            tCard!FontBold = 1
            tCard!FontUnderline = 0
            tCard!fontsize = 8
            tCard!TextAlign = taCenterTop
            tCard!Text = "„ﬂ »… √„Ì‰"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            ' ·Ì›Ê‰
            tCard.AddNew
            tCard!Top = MyMeasure(0.15) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!fontsize = 6
            tCard!TextAlign = taLeftTop
            tCard!Text = "Tel." & cHPhone1
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Top = MyMeasure(0.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!fontsize = 6
            tCard!TextAlign = taLeftTop
'            tCard!Text = cHPhone2
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.4) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(2)
            tCard!Height = MyMeasure(0.3)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!isBarcode = 1
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            tCard!Text = tPrint!Item
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' ITEM
            tCard.AddNew
            tCard!Top = MyMeasure(0.4) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1.1)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!FontUnderline = 1
            tCard!fontsize = 7
            tCard!TextAlign = taRightTop
                
            If Len(tPrint!Item) < 10 Then
                tCard!Text = tPrint!Item
            Else
'                tCard!Text = "/" & DelZero(tPrint!code) & "/"
            End If
                
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' DESCA
            tCard.AddNew
            tCard!Top = MyMeasure(0.67) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!fontsize = 7
            tCard!TextAlign = taCenterBottom
            tCard!Text = tPrint!desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


' PRICE
            tCard.AddNew
            tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!fontsize = 8
            tCard!TextAlign = taLeftTop
            If tPrint!isPRICE Then
                tCard!Text = "L.E. " & Format(tPrint!PRICE2, "Fixed")
                tCard!FontUnderline = 1
            Else
                tCard!Text = "L.E. " & Format(tPrint!price, "Fixed")
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
' COST
            If Val(tPrint!cost & "") > 0 Then
                tCard.AddNew
                tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
                tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
                tCard!Width = MyMeasure(1)
                tCard!Height = MyMeasure(0)
                tCard!FontName = "arial"
                tCard!FontBold = 0
                tCard!fontsize = 7
                tCard!TextAlign = taRightTop
                
                nCost = Val(LastPrice(tPrint!Item) & "")
                If nCost = 0 Then nCost = tPrint!cost
                If tPrint!isPRICE Then
                    nCost = (Val(tPrint!PRICE2 & "") - Val(tPrint!cost & "")) * 100 / 2
                Else
                    If Val(tPrint!package & "") <> 0 Then nCost = (Val(tPrint!price & "") - (nCost / Val(tPrint!package & ""))) / 2 * 100
                End If
                
                If IsNull(tPrint!doc_no) Then
                    tCard!Text = Format(nCost, "#0")
                Else
                    tCard!Text = Format(nCost, "#0") & "/" & DelZero(tPrint!doc_no)
                End If
                
                
                tCard!ForeColor = vbBlack
                tCard!CardNo = nCard
                tCard.Update
            End If
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 0 Then
        ItemTable.Find "item = " & MyParn(.EditText), , adSearchForward, adBookmarkFirst
        If Not ItemTable.EOF Then
            .TextMatrix(Row, 1) = ItemTable!desca & ""
            .TextMatrix(Row, 2) = ItemTable!price & ""
            .TextMatrix(Row, 3) = 1
        Else
            Cancel = 1
        End If
    End If
End With
End Sub

Private Sub xTotal56_Click()
Dim nTotal144 As Double
Dim nTotal56 As Double
With grid1
    For i = 1 To .Rows - 1
        If TurnValue(.TextMatrix(i, 4), "", False) Then
            If .TextMatrix(i, 5) = "1" Then
                nTotal56 = nTotal56 + Val(.TextMatrix(i, 3))
            Else
                nTotal144 = nTotal144 + Val(.TextMatrix(i, 3))
            End If
        End If
    Next i
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56
End Sub
Private Sub xTotal144_Click()
Dim nTotal144 As Double
Dim nTotal56 As Double
With grid1
    For i = 1 To .Rows - 1
        If TurnValue(.TextMatrix(i, 4), "", False) Then
            If .TextMatrix(i, 5) = "1" Then
                nTotal56 = nTotal56 + Val(.TextMatrix(i, 3))
            Else
                nTotal144 = nTotal144 + Val(.TextMatrix(i, 3))
            End If
        End If
    Next i
End With
xTotal144.Caption = nTotal144
xTotal56.Caption = nTotal56
End Sub

