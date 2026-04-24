VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salary_paid 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   11130
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   1365
      Begin VB.CommandButton cmdSave 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salary_paid.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "salary_paid.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   555
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   8205
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·»«ﬁÌ"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   945
         Width           =   450
      End
      Begin VB.Label xRest 
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   2130
      End
      Begin VB.Label xtotal 
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ì «·„— »"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label xPaid 
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ï «·„”œœ"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   1125
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "salary_paid.frx":48DC
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3645
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
      Height          =   2220
      Left            =   90
      TabIndex        =   0
      Top             =   1395
      Width           =   10950
      _cx             =   19315
      _cy             =   3916
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
End
Attribute VB_Name = "salary_paid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean, sDoc_no As String, SCODE As String
Dim con As New ADODB.Connection
Private Sub myloadgrd()
Dim cString
cString = "SELECT FILE2_51.DATE,FILE2_51.[VALUE],FILE2_51.DESCA,FILE2_51.ID FROM FILE2_51"
cString = cString & turn(cString) & "FILE2_51.DOC_NO = " & MyParn(sDoc_no)
cString = cString & turn(cString) & " FILE2_51.CODE = " & MyParn(SCODE)
DATA11.RecordSource = cString
DATA11.Refresh
MyAddItem
Fixgrd
Calctotals
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
Dim NROWS As Long
For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then NROWS = NROWS + 1
Next
If NROWS > 0 Then
    If myreplace Then
        Inform " „ «·Õ›Ÿ »‰Ã«Õ"
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
'If sboxSales = "" Then
'    MsgBox "·«  ÊÃœ Œ“‰… „»Ì⁄« "
'    Unload Me
'End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set charge_Cashfrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Calctotals

With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then MyAddItem
Calctotals
If myreplace(Row) Then
    If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadgrd
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub Calctotals()
Dim nTotal As Long
For I = 1 To grid1.Rows - 1
    nTotal = Val(grid1.TextMatrix(I, 1)) + nTotal
Next
xPaid.Caption = Myvalue(nTotal)
xRest.Caption = Val(xTotal.Caption) - Val(xPaid.Caption)
End Sub
Private Sub Grid1_EnterCell()
If bEdit Then
    If grid1.Col = 0 Then
        grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDKbdMouse
    End If
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Form_Load()
bEdit = True
openCon con
Dim aRet As Variant
aRet = GetFields("Select * FROM SALARY_BALANCE where doc_no = " & MyParn(sDoc_no) & " AND CODE = " & MyParn(SCODE))
If Not IsEmpty(aRet) Then
    xTotal.Caption = Val(retFlag(aRet, "TOTAL"))
End If
Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon
myloadgrd
'grid1.Select grid1.Rows - 1, 1
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then Cancel = True
ElseIf Col = 1 Then
    If Not ValidTotal(Row, grid1.EditText) Then
        MsgBox "«·„»·€ «·„”œœ «ﬂ»— „‰ «·«Ã„«·Ì «·„” Õﬁ"
        Cancel = True
    End If
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
        Calctotals
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, grid1.Cols - 1) = "" Then
    .RemoveItem .Row
    Calctotals
End If
End With
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
'If Not IsNumeric(.TextMatrix(Row, 3)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub Fixgrd()
Dim Col As Long, Row As Long
With grid1
.ColWidth(0) = 1500
.ColWidth(1) = 2000
.ColWidth(2) = 5000
.TextMatrix(0, 0) = "«· «—ÌŒ"
.TextMatrix(0, 1) = "«·ﬁÌ„…"
.TextMatrix(0, 2) = "«·»Ì«‰"
.ColHidden(.Cols - 1) = True
For Col = 0 To grid1.Cols - 1
    .ColAlignment(Col) = flexAlignRightCenter
Next

For Row = 1 To grid1.Rows - 1
    If grid1.TextMatrix(Row, .Cols - 1) <> "" Then grid1.Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = &H8000000F
Next
End With
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEdit Then
    If MsgBox("Õ–› «·”œœ ?", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from FILE2_51 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        Calctotals
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clea
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    If grid1.Col = 0 And IsDate(grid1.TextMatrix(grid1.Row, 0)) Then Exit Sub
    If grid1.Col = 1 And grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 1)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Function myreplace(Optional Row As Long = -1) As Boolean
With grid1
Dim aInsert As Variant
con.BeginTrans
On Error GoTo myerror
For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
    aInsert = AddFlag(Empty, "doc_no", addstring(sDoc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(grid1.TextMatrix(I, 0)))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(SCODE))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(.TextMatrix(I, 1)))
    aInsert = AddFlag(aInsert, "[desca]", addstring(.TextMatrix(I, 2)))
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
        con.Execute addInsert(aInsert, "FILE2_51")
    Else
        con.Execute addUpdate(aInsert, "FILE2_51", "ID = " & grid1.TextMatrix(Row, .Cols - 1))
    End If
Next
con.CommitTrans
End With
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub MyAddItem()
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = Format(Date, "YYYY-MM-DD")
End Sub
Private Function doprint() As Boolean
Dim nTotal As Double, NROWS As Long
On Error GoTo myerror
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
        nTotal = nTotal + Val(grid1.TextMatrix(I, 3))
        NROWS = NROWS + 1
    End If
Next

If NROWS = 0 Then
    doprint = True
    Exit Function
End If

With grid1
For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
        temptable.AddNew
        temptable!str1 = "«ﬁ—«— «” ·«„ ‰ﬁœÌ…"
        temptable!str2 = ArbString("«· «—ÌŒ : " & Format(dSalesDate, "yyyy/mm/dd"))
        temptable!str3 = ArbString("√ﬁ— «‰« «·”Ìœ/.................................................")
        temptable!str4 = "»√‰‰Ì «” ·„  „»·€ Êﬁœ—… : " & nTotal & " Ã‰ÌÂ"
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & MyOnly(nTotal)
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & "‰ŸÌ— :"
        temptable!str4 = ArbString(temptable!str4)
        temptable!str6 = TurnValue(.TextMatrix(I, 2))
        temptable!val2 = Val(.TextMatrix(I, 3))
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
Report1.Reset
FixPrinter Report1
Report1.ReportFileName = App.Path & "\reports\chargepaid.rpt"
Report1.DataFiles(0) = tempFile
Report1.Destination = crptToPrinter
Report1.Action = 1
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
Private Sub Handlecontrols()
cmdSave.Enabled = sboxSales <> ""
End Sub
Private Function ValidTotal(Row As Long, nValue As Double) As Boolean
Dim nTotal As Long
nTotal = nValue
For I = 0 To grid1.Rows - 1
    If I <> Row Then
        nTotal = nTotal + Val(grid1.TextMatrix(I, 1))
    End If
Next
If nTotal > Val(xTotal.Caption) Then Exit Function
ValidTotal = True
End Function
