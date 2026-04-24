VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Note_codesfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "√ﬂÊ«œ «·œ›« —"
   ClientHeight    =   6810
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   12120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   4410
      MaskColor       =   &H00FFFFFF&
      Picture         =   "note_codes.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5805
      UseMaskColor    =   -1  'True
      Width           =   1770
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1170
      Top             =   450
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
      Height          =   5685
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11895
      _cx             =   20981
      _cy             =   10028
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
      BackColorSel    =   8454143
      ForeColorSel    =   128
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
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
      WordWrap        =   -1  'True
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
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   6210
      TabIndex        =   1
      Top             =   5715
      Width           =   5775
      Begin VB.TextBox xBon_no 
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   4470
      End
      Begin VB.TextBox xBook_no 
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   4470
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·»Ê‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ «·œ› —"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   915
      End
   End
End
Attribute VB_Name = "Note_codesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Dim cList As String
Dim oSearchClient As New Search3, clist1 As String
Dim CardTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub myload()
On Error GoTo myerror
cString = "SELECT TYPE_GAS_CODES.CODE as [«·ﬂÊœ],TYPE_GAS_CODES.DESCA as [«·»Ì«‰],TYPE_GAS_CODES.PRICE AS [«·”⁄—],KIND AS [«·‰Ê⁄]" & _
          " FROM TYPE_GAS_CODES"
If Trim(xDesca.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "TYPE_GAS_CODES.desca")
End If
cString = cString & " order by TYPE_GAS_CODES.CODE"
data1.RecordSource = cString
data1.Refresh
myAddItem
FixGrd
Exit Sub
myerror:
 MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set TYPE_GAS_CODESfrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
If grid1.Row = grid1.Rows - 1 Then
    myAddItem
    grid1.ShowCell grid1.Rows - 1, 0
End If
If myreplace(Row) Then
   If grid1.TextMatrix(Row, 0) = "" Then myload
End If
End Sub
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
    If (grid1.Col = 0) Then
        grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDKbdMouse
    End If
End With
End Sub
Private Sub Grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub Form_Load()
openCon con
cList = StrList2("SELECT CODE,DESCA FROM KIND_GAS_CODES ORDER BY CODE")
Set grid1.DataSource = data1
data1.ConnectionString = strCon
With grid1

myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem OldRow
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim i As Long
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "—ﬁ„ «·œ› — €Ì— „”Ã·"
        Cancel = True
    Else
        Dim aRet As Variant, cString As String
        cString = "Select code from TYPE_GAS_CODES where desca = " & MyParn(grid1.EditText)
        If grid1.TextMatrix(Row, 0) <> "" Then cString = cString & turn(cString) & "code <> " & grid1.TextMatrix(Row, 0)
        aRet = GetField(cString)
        If Not IsEmpty(aRet) Then
            MsgBox "«·«”„ „ÊÃÊœ „‰ ﬁ»· ›Ï «·ﬂÊœ " & aRet
            Cancel = True
        End If
    End If
End If
End Sub
Private Sub FixGrd()
With grid1
.RowHeight(0) = 600
.ColWidth(0) = 900
.ColWidth(1) = 4000
.ColWidth(2) = 2500
.ColWidth(3) = 2500
.ColComboList(3) = cList
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xDesca_Change()
myload
End Sub
Private Function validRow(Row) As Boolean
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–› !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from TYPE_GAS_CODES where CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.Row
        End If
    End If
ElseIf KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
myload
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 3 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    If Col = 1 Then
        grid1.Select Row, NextEmpty(grid1, Row, Col + 1, 2)
    Else
        grid1.Select Row, Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 2)
    grid1.ShowCell Row + 1, 1
End If
End Sub
Private Sub myAddItem()
With grid1
    .AddItem ""
End With
End Sub
Private Function myreplace(Row As Long) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[DESCA]", addstring(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[PRICE]", Val(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[KIND]", addvalue(grid1.TextMatrix(Row, 3)))
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, 0) = "" Then
    Dim sCode As String
    sCode = Newflag("TYPE_GAS_CODES", "CODE", con)
    aInsert = AddFlag(aInsert, "[CODE]", sCode)
    con.Execute addInsert(aInsert, "TYPE_GAS_CODES")
Else
    con.Execute addUpdate(aInsert, "TYPE_GAS_CODES", "TYPE_GAS_CODES.CODE = " & grid1.TextMatrix(Row, 0))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Function

