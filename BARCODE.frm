VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form barcodefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»«— ﬂÊœ"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   12
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
   ScaleHeight     =   9660
   ScaleWidth      =   15270
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   9240
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1530
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   180
      Width           =   3570
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ÕœÌœ «·ﬂ·"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   6225
      Begin VB.CommandButton cmdBar56 
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1305
         Picture         =   "BARCODE.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelPrinted 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":7130
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4950
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":96A9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   90
      TabIndex        =   10
      Top             =   945
      Width           =   15135
      _cx             =   26696
      _cy             =   14473
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
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
      AutoResize      =   0   'False
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
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   3855
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄„Êœ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "«·’›:"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "barcodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemTable As New ADODB.Recordset
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Dim temptable As New ADODB.Recordset
Dim nCols As Double
Dim NROWS As Double


Private Sub cmdBar96_Click()

End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cString As String
cString = "update ADDPRINT SET ADDPRINT.ISPRINT = " & Check1.Value & " FROM ADDPRINT"
con.Execute cString
myloadgrd
End Sub

Private Sub CmdDelPrinted_Click()
    myDelete "Õ–› «·ﬂ·"
End Sub

Private Sub cmdExit_Click()
Unload Me
Set barcodefrm = Nothing
End Sub
Private Sub cmdSave_Click()
    If myreplaceGrd Then Inform " „ «·Õ›Ÿ"
    myloadgrd
End Sub
Private Sub cmduno_Click()
myloadgrd
End Sub
Private Sub CmdDel_Click()
myDelete "Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", "ISPRINT = 1"
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
'Grid1.LoadGrid App.Path & "\addPrint.grd", flexFileData
frmReturn.Show

End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "DELETE FROM ADDPRINT WHERE PRINT = 0"
    con.CommitTrans
    myloadgrd
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
myloadgrd
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myloadgrd
End Sub
Sub myloadgrd()
With grid1
 '                   0               1               2       3                   4                 5                   6               7             8               9           10              11              12
cString = "Select ADDPRINT.ITEM,FILE1_10.DESCA,ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM ADDPRINT  INNER JOIN FILE1_10 ON ADDPRINT.ITEM  = FILE1_10.ITEM "
data1.RecordSource = cString
data1.Refresh
MyAddItem
FixGrd
Calctotals
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim bNew As Boolean
If Col = 0 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
With grid1
If Row = .Rows - 1 Then
    MyAddItem
End If
Calctotals

If myreplaceGrd(Row) Then
   bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If
myloadgrd
If bNew Then
    'grid1.Row = grid1.Rows - 1
    grid1.ShowCell grid1.Rows - 1, 1
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 2 Or grid1.Col = 5 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then .RemoveItem OldRow
End If
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean) As Boolean
With grid1
If (.TextMatrix(Row, 0)) = "" Then Exit Function
If (.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub delcheck()
For i = 1 To grid1.Rows - 2
   If Val(grid1.TextMatrix(i, 8)) <> 0 Then
        grid1.RowHidden(i) = True
   End If
Next
myreplaceGrd
End Sub
Sub myProc()
Dim bNew As Boolean
bNew = grid1.Row = grid1.Rows - 1
grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Grid1_AfterEdit grid1.Row, grid1.Col
If Not bNew Then
    Unload oSearchItem
    CellPos 13, grid1.Row, grid1.Col
Else
    grid1.Row = grid1.Rows - 1
    grid1.Col = 0
End If
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
Private Sub FixGrd()
With grid1
    .FormatString = "»«—ﬂÊœ|" & "«·«”„|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "—ﬁ„ «·„” ‰œ|" & "«·ÿ»«⁄…|"
    .ColWidth(0) = 2000
    .ColWidth(1) = 5000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColHidden(4) = True
    .ColDataType(5) = flexDTBoolean
    .ColHidden(.Cols - 1) = True
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Function myreplaceGrd(Optional nRow As Long = -1) As Boolean
Dim aInsert(3, 1)
With grid1
For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
     aInsert(0, 0) = "doc_no"
     aInsert(0, 1) = addstring(grid1.TextMatrix(i, 4))
     
     aInsert(1, 0) = "item"
     aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
             
     aInsert(2, 0) = "quant"
     aInsert(2, 1) = Val(.TextMatrix(i, 2))
    
     aInsert(3, 0) = "isprint"
     aInsert(3, 1) = IIf(Val(.TextMatrix(i, 5)) <> 0, 1, 0)
     
     If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
         con.Execute CreateInsert(aInsert, "ADDPRINT")
     Else
         con.Execute CreateUpdate(aInsert, "ADDPRINT", " where ID = " & grid1.TextMatrix(i, .Cols - 1), Array(-1))
     End If
Next
End With
myreplaceGrd = True
End Function
'Private Sub GrdDesc(ByVal Row As Long, Col As Long)
'Dim loctable  As ADODB.Recordset
'If IsNumeric(grid1.TextMatrix(Row, 2)) Then
'    Set loctable = RetItemRow(grid1.TextMatrix(Row, 2), con)
'    If Not (loctable.EOF And loctable.BOF) Then
'        grid1.TextMatrix(Row, 1) = loctable!Item & ""
'        grid1.TextMatrix(Row, 3) = loctable!Model & ""
'
'        Dim ModelTable As New ADODB.Recordset
'        Set ModelTable = ModelByModel(loctable!Model & "", con, "ModelByModelAll")
'        If Not (ModelTable.EOF And ModelTable.BOF) Then
'            grid1.TextMatrix(Row, 4) = ModelTable!X1 & ""
'        End If
'
'        Set ModelTable = Nothing
'        grid1.TextMatrix(Row, 5) = loctable!SCAL
'        grid1.TextMatrix(Row, 6) = loctable!Color & ""
'        grid1.TextMatrix(Row, 7) = "1"
'        grid1.TextMatrix(Row, 8) = Format(loctable!PRICE, "#0.00")
'        grid1.TextMatrix(Row, 9) = Format(loctable!Price2, "#0.00")
'        grid1.TextMatrix(Row, 11) = -1
'    End If
'End If
'Set loctable = Nothing
'End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        Calctotals
        'makeSerial Grid1.Row
    End If
ElseIf KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Function RemoveItem(Row As Long) As Boolean
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From ADDPRINT where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
End If
grid1.RemoveItem Row
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub myDelete(cMsg As String, Optional cFilter As String)
If MsgBox(cMsg, vbYesNo + vbDefaultButton2) = vbYes Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "DELETE FROM ADDPRINT FROM ADDPRINT" & turn(cFilter) & cFilter
    
    con.CommitTrans
    myloadgrd
End If
Exit Sub
myerror:
    con.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 0 Then
        '.TextMatrix(Row, Col) = .EditText
        'GrdDesc (Row)
        If Trim(.EditText) = "" Then
            Cancel = True
            '.TextMatrix(Row, Col) = ""
        End If
    End If
End With
End Sub
Private Sub GrdDesc(Row)
With grid1
'cString = "Select 0 ADDPRINT.ITEM,1FILE1_10.DESCA,2ADDPRINT.QUANT,3FILE1_10.PRICE,4ADDPRINT.DOC_NO,5ADDPRINT.ISPRINT,6ADDPRINT.ID "
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 3) = ""

If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Sub
Dim aRet As Variant
aRet = aGetDesca("SELECT FILE1_10.DESCA,FILE1_10.PRICE FROM FILE1_10 WHERE FILE1_10.ITEM = " & MyParn(grid1.TextMatrix(Row, 0)))
If UBound(aRet) > 0 Then
    grid1.TextMatrix(Row, 1) = aRet(1)
    grid1.TextMatrix(Row, 3) = aRet(2) & ""
    If Trim(grid1.TextMatrix(Row, 2)) = "" Then grid1.TextMatrix(Row, 2) = 1
End If
End With
End Sub
Private Sub cmdBar56_Click()
If Val(xRow.Text) > 14 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 4 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If

myreplaceGrd
DoprintStr56
Set myForm = Me
CardPrintNew.Show 1
myDelete "Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", "ISPRINT = 1"
End Sub
Private Sub MyAddItem()
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 5) = -1
'MakeSerial
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col = 0 Then
    grid1.Col = 2
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = grid1.Row + 1
    grid1.Col = 0
    grid1.ShowCell Row + 1, 0
End If
End Sub
Private Sub DoprintStr56()
Dim tCard As New ADODB.Recordset
Dim tPrint As New ADODB.Recordset
Dim nCost As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(0.4)
SettingArray(cCardWidth) = MyMeasure(5.25)
SettingArray(cCardHeight) = MyMeasure(2.1)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 14
SettingArray(cCols) = 4
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete From Card"
tCard.Open "Select * From card", con, adOpenKeyset, adLockOptimistic, adCmdText
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
        tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
        tCard!Width = MyMeasure(3.4)
        tCard!Height = MyMeasure(0)
        tCard!FontName = "Times New Roman"
        tCard!FontBold = True
        tCard!FontUnderline = 0
        tCard!fontsize = 14
        tCard!TextAlign = taCenterTop
        tCard!Text = "Ali's Market"
        tCard!ForeColor = vbBlack
        tCard!CardNo = nCard
        tCard.Update
            
            
' BARCODE
            tCard.AddNew
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.4)
            tCard!Height = MyMeasure(0.5)
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
            tCard!FontName = "Times New Roman"
            tCard!FontBold = 1
            tCard!fontsize = 8
            tCard!TextAlign = taRightTop
            tCard!Text = tPrint!Item
            tCard!FontUnderline = True

            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            
' DESCA
            tCard.AddNew
            tCard!Top = MyMeasure(1.4) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(5.25)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Times New Roman"
            tCard!FontBold = False
            tCard!fontsize = 9
            tCard!TextAlign = taCenterBottom
            tCard!Text = tPrint!Desca
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


' PRICE
            tCard.AddNew
            tCard!Top = MyMeasure(1.7) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.4) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "Times New Roman"
            tCard!FontBold = True
            tCard!fontsize = 10
            tCard!TextAlign = taLeftTop
'            tCard!Text = Format(tPrint!price, "Fixed") & " LE"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            nHeight = SettingArray(cCardHeight)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Sub xitem_GotFocus()
myGotFocus XITEM
End Sub

Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Trim(XITEM.Text) = "" Then Exit Sub
    Dim cString As String, nAffect As Long
    cString = "update ADDPRINT SET ADDPRINT.ISPRINT = 1 FROM ADDPRINT"
    cString = cString & turn(cString) & " ADDPRINT.ITEM = " & MyParn(XITEM.Text)
    con.Execute cString, nAffect
    If nAffect > 0 Then
        Inform " „ «÷«›… «·ÿ»«⁄… ··»«—ﬂÊœ"
        XITEM.Text = ""
    Else
        MsgBox "«·»«—ﬂÊœ €Ì— ’ÕÌÕ «Ê ·Ì” Â‰«ﬂ »«—ﬂÊœ ··’‰› «·„Õœœ"
        myGotFocus XITEM
    End If
    myloadgrd
    XITEM.SetFocus
End If
End Sub
Private Sub Calctotals()
Dim nCard As Long, nPage As Long, nRest As Long
With grid1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 5)) <> 0 Then
            nTotal = nTotal + Val(Val(.TextMatrix(i, 2)))
        End If
    Next
    nPage = Int(nTotal / 56) + IIf(Int(nTotal / 56) < (nTotal / 56), 1, 0)
    nRest = 56 - (nTotal Mod 56)
End With
StatusBar1.Panels(1).Text = "⁄œœ «·ﬁÿ⁄ : " & nTotal
StatusBar1.Panels(2).Text = "⁄œœ «·’›Õ«  : " & nPage
StatusBar1.Panels(3).Text = "ﬁÿ⁄ »«ﬁÌ… : " & (nPage * 56) - nTotal
'StatusBar1.Panels(3).Text = "’›Ê› €Ì— „ÿ»Ê⁄… : " & Int(nRest / 14) + IIf(Int(nRest / 14) < (nRest / 14), 1, 0)
End Sub
