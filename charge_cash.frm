VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form charge_Cashfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   11130
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   1365
      Begin VB.CommandButton cmdSave 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge_cash.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge_cash.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   555
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   8205
      Begin VB.CheckBox chkPrint 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·€«¡ «·ÿ»«⁄…"
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
         Height          =   285
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   675
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   5505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ«‘Ì— :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "charge_cash.frx":48DC
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   5220
      UseMaskColor    =   -1  'True
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -2880
      Top             =   -45
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Height          =   4110
      Left            =   135
      TabIndex        =   0
      Top             =   1035
      Width           =   10950
      _cx             =   19315
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   -3105
      Top             =   -180
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label xtotal 
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
      Height          =   375
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5265
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ì «·„’—Ê› :"
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
      Left            =   9405
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5310
      Width           =   1545
   End
End
Attribute VB_Name = "charge_Cashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim oSearchCode As New Search3
Private Sub myloadgrd()
Dim cString
cString = "SELECT CHARGE,FILE8_51.DESCA,FILE8_50.DESCA,FILE8_50.VALUE,FILE8_50.ID FROM FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO LEFT JOIN FILE8_51 ON FILE8_50.CHARGE = FILE8_51.CODE"
'cString = cString & turn(cString) & "FILE8_50H.DOC_NO_CASH = " & MyParn(Format(dSalesDate, "YYMMDD") & sboxSales)
cString = cString & turn(cString) & "FILE8_50H.DATE = " & DateSq(Format(dSalesDate, "DD-MM-YYYY"))
cString = cString & turn(cString) & "FILE8_50.BOX = " & MyParn(sboxSales)
data11.RecordSource = cString
data11.Refresh
Fixgrd
myAddItem
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim NROWS As Long
For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then NROWS = NROWS + 1
Next
If NROWS > 0 Then
    If myreplace Then
        Inform " „ «·Õ›Ÿ »‰Ã«Õ"
        If Me.chkPrint.Value = 0 Then
            If Not doprint Then MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ «·ÿ»«⁄…!!"
        End If
        myloadgrd
        Unload Me
    Else
        If MsgBox("·„ Ì „ «·Õ›Ÿ !!  —«Ã⁄ ⁄‰ «·Œ—ÊÃ", vbOKCancel) <> vbOK Then Unload Me
    End If
Else
    Unload Me
End If
End Sub

Private Sub CmdUndo_Click()
myloadgrd
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
End Sub

Private Sub Form_Activate()
If sboxSales = "" Then
    MsgBox "·«  ÊÃœ Œ“‰… „»Ì⁄« "
    Unload Me
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
addSetting "print", chkPrint.Value, TempSave(Me)
closeCon con
Set charge_Cashfrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then GrdDesc Row
If validRow(Row) And Row = grid1.Rows - 1 Then myAddItem
CalcTotals
End Sub
Private Sub CalcTotals()
Dim nTotal As Long
For i = 1 To grid1.Rows - 1
    nTotal = Val(grid1.TextMatrix(i, 3)) + nTotal
Next
xtotal.Caption = Myvalue(nTotal)
End Sub
Private Sub grid1_EnterCell()
If bedit Then
    grid1.Editable = IIf(grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Or grid1.Col = 1, flexEDNone, flexEDKbdMouse)
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) = "" Then
    If MsgBox("Õ–› !! Â· «‰  „Ê«›ﬁ ø", vbYesNo) = vbYes Then
        '        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        '            con.BeginTrans
        '            On Error GoTo myerror
        '            Dim cString As String
        '            cString = "Delete from resale_codes"
        '            cString = cString & turn(cString) & " ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        '            con.Execute cString
        '            con.CommitTrans
        '        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Form_Load()
xDate.Caption = Format(dSalesDate, "DD-MM-YYYY")
If sboxSales <> "" Then xDesca.Caption = GetDesca("Select desca from file0_50 where code = " & MyParn(sboxSales))
chkPrint.Value = Val(RetSetting("print", TempSave(Me)))
bedit = True
openCon con

Set grid1.DataSource = data11
data11.ConnectionString = strCon
myloadgrd
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then Cancel = True
ElseIf Col = 3 Then
   If Not IsNumeric(grid1.EditText) Then
        MsgBox "Numeric Value Requiered"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, grid1.Cols - 1) = "" Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 3)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub Fixgrd()
Dim Col As Long, Row As Long
With grid1
.ColWidth(0) = 700
.ColWidth(1) = 2000
.ColWidth(2) = 5000
.ColWidth(3) = 1000
.TextMatrix(0, 0) = "«·ﬂÊœ"
.TextMatrix(0, 1) = "«·„’—Ê›"
.TextMatrix(0, 2) = "«·»Ì«‰"
.TextMatrix(0, 3) = "«·ﬁÌ„…"
.ColHidden(.Cols - 1) = True
For Col = 0 To grid1.Cols - 1
    .ColAlignment(Col) = flexAlignRightCenter
Next

For Row = 1 To grid1.Rows - 1
    If grid1.TextMatrix(Row, .Cols - 1) <> "" Then grid1.Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &H8000000F
Next
End With
End Sub
Private Sub xStatus_Change()
myloadgrd
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    grdLookup
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1 + IIf(Col = 0, 1, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, 0
    grid1.ShowCell Row + 1, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myreplaceGrd(sDoc_No As String)
With grid1
Dim aInsert(4, 1)
For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, .Cols - 1) = "" Then
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(sDoc_No)
        
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(sboxSales)
        
        aInsert(2, 0) = "Charge"
        aInsert(2, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(3, 0) = "DESCA"
        aInsert(3, 1) = addstring(.TextMatrix(i, 2))
        
        aInsert(4, 0) = "VALUE"
        aInsert(4, 1) = Val(.TextMatrix(i, 3))
        con.Execute CreateInsert(aInsert, "FILE8_50")
    End If
Next
End With
End Sub
Private Function myreplace() As Boolean
Dim aInsert(2, 1), sDoc_No As String, cString As String
cString = "Select doc_no from  file8_50H"
cString = cString & turn(cString) & "DOC_NO_CASH = " & MyParn(Format(dSalesDate, "YYMMDD") & sboxSales)
sDoc_No = GetDesca(cString)

aInsert(0, 0) = "DOC_NO"

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(dSalesDate)

aInsert(2, 0) = "Doc_No_cash"
aInsert(2, 1) = addstring(Format(dSalesDate, "YYMMDD") & sboxSales)

con.BeginTrans
On Error GoTo myerror
If sDoc_No = "" Then
    sDoc_No = RetZero(Val(Newflag("FILE8_50H", "doc_no")))
    aInsert(0, 1) = addstring(sDoc_No)
    con.Execute CreateInsert(aInsert, "FILE8_50H")
Else
    aInsert(0, 1) = addstring(sDoc_No)
    con.Execute CreateUpdate(aInsert, "FILE8_50H", " where doc_no = " & addstring(sDoc_No))
End If
myreplaceGrd sDoc_No
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 1) = ""
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Sub
grid1.TextMatrix(Row, 0) = RetZero(grid1.TextMatrix(Row, 0), 3)
grid1.TextMatrix(Row, 1) = GetDesca("Select Desca From file8_51 Where code = " & MyParn(grid1.TextMatrix(Row, 0))) & ""
End Sub
Sub myProc()
grid1.TextMatrix(grid1.Row, 0) = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
grid1.TextMatrix(grid1.Row, 1) = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 1)
grid1_AfterEdit grid1.Row, grid1.Col
CellPos 13, grid1.Row, grid1.Col
Unload oSearchCode
End Sub
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE8_51"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCode.Caption = "≈” ⁄·«„ "
oSearchCode.Show 1
End Sub
Private Sub myAddItem()
grid1.AddItem ""
End Sub
Private Function doprint() As Boolean
Dim nTotal As Double, NROWS As Long
On Error GoTo myerror
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
        nTotal = nTotal + Val(grid1.TextMatrix(i, 3))
        NROWS = NROWS + 1
    End If
Next

If NROWS = 0 Then
    doprint = True
    Exit Function
End If

With grid1
For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
        temptable.AddNew
        temptable!str1 = "«ﬁ—«— «” ·«„ ‰ﬁœÌ…"
        temptable!str2 = ArbString("«· «—ÌŒ : " & Format(dSalesDate, "yyyy/mm/dd"))
        temptable!str3 = ArbString("√ﬁ— «‰« «·”Ìœ/.................................................")
        temptable!str4 = "»√‰‰Ì «” ·„  „»·€ Êﬁœ—… : " & nTotal & " Ã‰ÌÂ"
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & MyOnly(nTotal)
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & "‰ŸÌ— :"
        temptable!str4 = ArbString(temptable!str4)

        'temptable!str6 = TurnValue(ArbString(.Cell(flexcpTextDisplay, i, 0, i, 0) & turn(.TextMatrix(i, 2), "[" & .TextMatrix(i, 2) & "]")))
        temptable!str6 = TurnValue(.TextMatrix(i, 2))
        temptable!val2 = Val(.TextMatrix(i, 3))
        temptable.Update
    End If
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    doprint = True
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
REPORT1.Reset
FixPrinter REPORT1
REPORT1.ReportFileName = App.Path & "\reports\chargepaid.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.Destination = crptToPrinter
REPORT1.Action = 1
closeCon:
temptable.Close
Set temptable = Nothing
doprint = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
'GoTo closeCon
End Function

