VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form charge_codefrm 
   Caption         =   "«ﬂÊ«œ «·„’«—Ì›"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   11490
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   3825
      MaskColor       =   &H00FFFFFF&
      Picture         =   "charge_code.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4815
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   5310
      TabIndex        =   1
      Top             =   4725
      Width           =   6090
      Begin VB.TextBox xDesca 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   4110
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   330
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   345
         Left            =   450
         TabIndex        =   4
         Top             =   180
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label label2 
         Caption         =   "«·»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "«·„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1770
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc DATA3 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4650
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11310
      _cx             =   19950
      _cy             =   8202
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   12
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
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
End
Attribute VB_Name = "charge_codefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cList As String
Private Sub myload()
Dim cString As String
cString = "SELECT CODE as [«·ﬂÊœ], DESCA as [«·»Ì«‰],[GROUP] as [«·„Ã„Ê⁄…],MONTHES AS [«·„œ… »«·‘Â—],KILOS AS [ﬂÌ·Ê „ —]" & _
          " FROM FILE8_51"
If xGroup.MatchedWithList Then cString = cString & turn(cString) & "[GROUP] = " & MyParn(xGroup.BoundText)
If Trim(xDesca.Text) <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "desca")
cString = cString & " order by CODE"
Set data1.Recordset = myRecordSet(cString, con)
myAddItem
FixGrd
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGroup_Click()
Dim oFlagGroup As New FlagGroupFrm, sSave As String
sSave = xGroup.BoundText
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sCaption = "„Ã„Ê⁄… «·„’—Ê›"
oFlagfrm.sFieldCode = "[CODE]"
oFlagfrm.sFieldDesca = "[DESCA]"
oFlagfrm.sTable = "FILE8_52"
oFlagfrm.sFieldName1 = "«·ﬂÊœ"
oFlagfrm.sFieldName2 = "«·»Ì«‰"
oFlagfrm.nZero = nZeroGroup
oFlagfrm.bedit = True
oFlagfrm.Show 1
DATA2.Refresh
xGroup.BoundText = sSave
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set FILE8_51frm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
    myAddItem
    grid1.ShowCell grid1.Rows - 1, 0
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DESCA]", addstring(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[GROUP]", addstring(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[MONTHES]", Val(grid1.TextMatrix(Row, 3)))
aInsert = AddFlag(aInsert, "[KILOS]", Val(grid1.TextMatrix(Row, 4)))
On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(Row, 0) = "" Then
    grid1.TextMatrix(Row, 0) = RetZero(Newflag("FILE8_51", "CODE"), 3)
    aInsert = AddFlag(aInsert, "[CODE]", addstring(grid1.TextMatrix(Row, 0)))
    con.Execute addInsert(aInsert, "FILE8_51")
Else
    con.Execute addUpdate(aInsert, "FILE8_51", "CODE = " & MyParn(grid1.TextMatrix(Row, 0)))
End If
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub Grid1_EnterCell()
If (grid1.Col = 0) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Form_Load()
openCon con
cList = StrList("select code,desca from file8_52 order by desca")

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE8_52"
Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set grid1.DataSource = data1
With grid1
myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End With
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row > 0 Then
    If MsgBox("Õ–› !! Â· «‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2) = vbYes Then
        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "Delete from FILE8_51 where code = " & grid1.TextMatrix(grid1.Row, 0)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myload
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 2 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "»Ì«‰ «·„’—Ê› €Ì— „”Ã·"
        Cancel = True
    Else
        Dim aRet As Variant
        aRet = GetField("Select code from file8_51 where desca = " & MyParn(grid1.EditText) & " and code <> " & MyParn(grid1.TextMatrix(Row, 0)))
        If Not IsEmpty(aRet) Then
            MsgBox "«·«”„ „ÊÃÊœ „‰ ﬁ»· ›Ï «·ﬂÊœ " & aRet
            Cancel = True
        End If
    End If
End If
End Sub
Private Sub FixGrd()
With grid1
.ColWidth(0) = 800
.ColWidth(1) = 3000
.ColWidth(2) = 3000
.ColWidth(3) = 2000
.ColWidth(4) = 2000
.ColComboList(2) = cList
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xDesca_Change()
myload
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, 0) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .Rows - 1 And .Row <> 0 And .TextMatrix(grid1.Row, 0) = "" Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(Row) As Boolean
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
'If Trim(grid1.TextMatrix(Row, 2)) = "" Then Exit Function
validRow = True
End Function
Private Sub myAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 2)
    grid1.ShowCell grid1.Row, 0
End If
End Sub
Private Sub xGroup_Change()
If xGroup.MatchedWithList Or Trim(xGroup.BoundText) = "" Then
    myload
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub

Private Sub xGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    grid1.SetFocus
    CellPos KeyCode, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub

Private Sub xGroup_LostFocus()
If Not xGroup.MatchedWithList Then
    xGroup.BoundText = ""
    myload
End If
End Sub
