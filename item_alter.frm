VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form item_alterfrm 
   Caption         =   "«ﬂÊ«œ »œÌ·…"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDel 
      Caption         =   "Õ–› »œ«∆· »œÊ‰ «’‰«›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1035
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   5775
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   210
         Width           =   3030
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   " €ÌÌ— «·»œÌ·"
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   4380
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   960
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -1305
      Top             =   -45
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   1530
      Width           =   7935
      _cx             =   13996
      _cy             =   7250
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Cols            =   7
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
      Editable        =   2
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
End
Attribute VB_Name = "item_alterfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Public sItem As String, sDesca As String

Private Sub cmdChange_Click()
If GetDesca("Select item from file1_10 where item = " & MyParn(xItem.Text)) = "" Then
    MsgBox "·« ÌÊÃœ ’‰› »Â–« «·ﬂÊœ"
    Exit Sub
End If
Dim cString As String
cString = "UPDATE ITEM_ALTER SET ITEM_ALTER.ITEM = " & addstring(xItem.Text)
cString = cString & turn(cString) & "ITEM_ALTER.ITEM = " & addstring(xItem.Tag)
con.BeginTrans
'On Error GoTo myerror
con.Execute cString, nRecord
con.CommitTrans
If nRecord > 0 Then Inform " „  ⁄œÌ· «·ﬂÊœ »‰Ã«Õ"
xItem.Tag = xItem.Text
myloadgrd
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
xItem.Text = xItem.Tag
End Sub
Private Sub CmdDel_Click()
Dim cString As String, nRecord As Long
cString = "DELETE FROM ITEM_ALTER WHERE ITEM NOT IN(SELECT ITEM FROM FILE1_10)"
con.BeginTrans
On Error GoTo myerror
con.Execute cString, nRecord
con.CommitTrans
Inform " „ Õ–› " & nRecord & "»œ«∆· »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Load()
openCon con
xItem.Tag = sItem
xItem.Text = sItem
cmdChange.Enabled = False
xDesca.Caption = sDesca
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myloadgrd
grid1.ShowCell grid1.Rows - 1, 1
grid1.Select grid1.Rows - 1, 1
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
myreplace Row, Col
End Sub
Private Sub FixSerial()
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT ITEM_ALTER.ITEMSUB,DESCA, ITEM_ALTER.ID " & _
          " FROM ITEM_ALTER "
    cString = cString & turn(cString) & " ITEM_ALTER.ITEM = " & MyParn(xItem.Text)
    cString = cString & "ORDER by ITEM_ALTER.ROW"
    data1.RecordSource = cString
    data1.Refresh
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = xDesca.Caption
    FixSerial
End With
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„|ﬂÊœ «·’‰›|≈”„ «·’‰›|"
.ColWidth(0) = 400
.ColWidth(1) = 2000
.ColWidth(2) = 4000
.ColHidden(.Cols - 1) = True
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

End With
End Sub
Private Sub myreplace(Row, Col)
Dim nBalance As Double

If Not validRows(Row) Then Exit Sub

If Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(i, 2) = xDesca.Caption
    FixSerial
End If
myreplaceGrdRow Row
grid1.Row = grid1.Rows - 1: grid1.Col = 1
End Sub
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    If Not validRows(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 2 Then
    grid1.EditCell
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        FixSerial
    End If
End If
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
If Not validRows(grid1.Row) And grid1.Row <> grid1.Rows - 1 Then grid1.RemoveItem grid1.Row
End Sub
Private Function myreplaceGrdRow(i) As Boolean
Dim aInsert(3, 1)
With grid1
con.BeginTrans
aInsert(0, 0) = "ITEM"
aInsert(0, 1) = addstring(xItem.Text)

aInsert(1, 0) = "ITEMSUB"
aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))

aInsert(2, 0) = "DESCA"
aInsert(2, 1) = addstring(grid1.TextMatrix(i, 2))

aInsert(3, 0) = "row"
aInsert(3, 1) = i
On Error GoTo myerror
If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, "ITEM_ALTER")
Else
    con.Execute CreateUpdate(aInsert, "ITEM_ALTER", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
End If
End With
con.CommitTrans
If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then myloadgrd
myreplaceGrdRow = True
Exit Function
myerror:
con.RollbackTrans
If Err.Number = -2147217873 Then
   MsgBox "«·’‰› „ÊÃÊœ ﬂ»œÌ· " & GetDesca("SELECT FILE1_10.DESCA FROM FILE1_10 INNER JOIN ITEM_ALTER ON FILE1_10.ITEM = ITEM_ALTER.ITEM WHERE ITEMSUB = " & MyParn(grid1.TextMatrix(i, 1)))
Else
    MsgBox Err.Description
End If
Err.Clear
myloadgrd
End Function
Private Function validRows(Optional prow = -1, Optional igMsg As Boolean = True, Optional bReqQuant As Boolean = False) As Boolean
For nRow = IIf(prow = -1, 1, prow) To IIf(prow = -1, grid1.Rows - 2, prow)
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then
        If Not igMsg Then MsgBox "«·’‰› ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã· "
        Exit Function
    End If
Next
validRows = True
End Function
Private Function RemoveItem(nRow) As Boolean
On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From ITEM_ALTER where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
End If
con.CommitTrans
grid1.RemoveItem nRow
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub xITEM_Change()
If Trim(xItem.Text) = "" Then Exit Sub
If Trim(LCase(xItem.Text)) <> Trim(LCase(xItem.Tag)) Then cmdChange.Enabled = True
End Sub
Private Sub xitem_Validate(Cancel As Boolean)
Dim sDesca As String
xDesca.Caption = ""
sDesca = GetDesca("Select desca from file1_10 where item = " & MyParn(xItem.Text))
If Trim(sDesca) = "" Then Exit Sub
xDesca.Caption = sDesca
myloadgrd
End Sub
