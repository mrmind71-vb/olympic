VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form debit_codesfrm 
   Caption         =   "«ﬂÊ«œ «·œ«∆‰Ê‰"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   15225
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2295
      Top             =   315
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
      Left            =   4455
      Top             =   1575
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
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   15135
      _cx             =   26696
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
      Cols            =   7
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
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   10440
      TabIndex        =   4
      Top             =   5760
      Width           =   4740
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   330
      End
      Begin VB.TextBox xDesca 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   3390
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   450
         TabIndex        =   2
         Top             =   540
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ :"
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   8910
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5805
      Width           =   1500
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "debit_codes.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
   End
End
Attribute VB_Name = "debit_codesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Dim oSearchClient As New Search3
Dim CardTable As New ADODB.Recordset, clist1 As String
Dim con As New ADODB.Connection
Private Sub myload()
On Error GoTo myerror
cString = "SELECT FILE8_81.CODE as [«·ﬂÊœ],FILE8_81.DESCA as [«·«”„],FILE8_81.[GROUP] AS [«·„Ã„Ê⁄…],CONVERT(VARCHAR(10),FILE8_81.DATE_FIRST,111) AS [ «—ÌŒ »œ«Ì… «· ⁄«„·],FILE8_81.FIRST_BAL AS [—’Ìœ «·»œ«Ì…],CONVERT(VARCHAR(10),FILE8_81.DATE_LAST,111) AS [ «—ÌŒ ‰Â«Ì… «· ⁄«„·], FILE8_81.CLIENT AS [ﬂÊœ «·⁄„Ì·],FILE3_10.DESCA AS [≈”„ «·⁄„Ì·],FILE8_81.CODE" & _
          " FROM FILE8_81 LEFT JOIN FILE3_10 ON FILE8_81.CLIENT = FILE3_10.CODE"
If Trim(xDesca.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "FILE8_81.desca")
End If
If xGroup.MatchedWithList Then
    cString = cString & turn(cString) & "FILE8_81.[GROUP] = " & xGroup.BoundText
End If
cString = cString & " order by FILE8_81.CODE"
data1.RecordSource = cString
data1.Refresh
myAddItem
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGroup_Click()
Dim oFlagfrm As New flag_mainfrm, cValue As String
cValue = xGroup.BoundText
oFlagfrm.sTable = "FILE8_81G"
oFlagfrm.sCaption = "„Ã„Ê⁄… «·œ«∆‰Ê‰"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE8_81G", con)
xGroup.BoundText = cValue
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
Fixgrd
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set debit_codesfrm = Nothing
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

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, 0) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub Grid1_EnterCell()
With grid1
    If (grid1.Col = 0 Or grid1.Col = 7) Then
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
clist1 = StrList("Select code,desca from file8_81G order by desca")
openCon con

Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE8_81G", con)
Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

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
        MsgBox "«·«”„ €Ì— „”Ã·"
        Cancel = True
    Else
        Dim aRet As Variant
        aRet = GetField("Select code from file8_81 where desca = " & MyParn(grid1.EditText) & " and code <> " & MyParn(grid1.TextMatrix(Row, 0)))
        If Not IsEmpty(aRet) Then
            MsgBox "«·«”„ „ÊÃÊœ „‰ ﬁ»· ›Ï «·ﬂÊœ " & aRet
            Cancel = True
        End If
    End If
ElseIf Col = 3 Then
    grid1.EditText = Format(grid1.EditText, "YYYY/MM/DD")
    If Not IsDate(grid1.EditText) Then
        MsgBox IIf(Trim(grid1.EditText) = "", "«· «—ÌŒ €Ì— „”Ã·", "«· «—ÌŒ €Ì— ”·Ì„")
        Cancel = True
    End If
ElseIf Col = 5 Then
    grid1.EditText = Format(grid1.EditText, "YYYY/MM/DD")
    If Trim(grid1.EditText) <> "" And Not IsDate(grid1.EditText) Then
        MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
        Cancel = True
    End If
ElseIf Col = 6 Then
    If Trim(grid1.EditText) = "" Then
        grid1.TextMatrix(Row, 7) = ""
    Else
        grid1.EditText = RetZero(grid1.EditText)
        aRet = GetFields("SELECT CODE,DESCA FROM FILE3_10 WHERE CODE = " & MyParn(grid1.EditText), con)
        If IsEmpty(aRet) Then
           MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
           Cancel = True
        Else
            grid1.TextMatrix(Row, 7) = retFlag(aRet, "DESCA") & ""
        End If
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.RowHeight(0) = 500
.ColWidth(0) = 900
.ColWidth(1) = 3500
.ColWidth(2) = 2000
.ColWidth(3) = 1400
.ColWidth(4) = 1200
.ColWidth(5) = 1400
.ColWidth(6) = 900
.ColWidth(7) = 3000
.ColComboList(2) = clist1
.ColHidden(.Cols - 1) = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xDesca_Change()
myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Function validRow(Row) As Boolean
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
If Not IsDate(grid1.TextMatrix(Row, 3)) Then Exit Function
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 6 Then
    ClientLookupAll Me, oSearchClient
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–› !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from FILE8_81 where CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
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
If KeyCode = 13 And Col <> 2 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 3 Then
    If Col = 1 Then
        grid1.Select Row, NextEmpty(grid1, Row, Col + 1, 3)
    Else
        grid1.Select Row, Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 3)
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
aInsert = AddFlag(aInsert, "[GROUP]", addvalue(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[DATE_FIRST]", addDate(grid1.TextMatrix(Row, 3)))
aInsert = AddFlag(aInsert, "[FIRST_BAL]", Val(grid1.TextMatrix(Row, 4)))
aInsert = AddFlag(aInsert, "[DATE_LAST]", addDate(grid1.TextMatrix(Row, 5)))
aInsert = AddFlag(aInsert, "[CLIENT]", addstring(grid1.TextMatrix(Row, 6)))
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    Dim sCode As String
    sCode = RetZero(Newflag("FILE8_81", "CODE", con), 3)
    aInsert = AddFlag(aInsert, "[CODE]", addstring(sCode))
    con.Execute addInsert(aInsert, "FILE8_81")
Else
    con.Execute addUpdate(aInsert, "FILE8_81", "FILE8_81.CODE = " & MyParn(grid1.TextMatrix(Row, 0)))
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
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 6 Then
        grid1.TextMatrix(grid1.Row, 6) = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 7) = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 1)
        Grid1_AfterEdit grid1.Row, grid1.Col
        Unload oSearchClient
        CellPos 13, grid1.Row, grid1.Col
    End If
End If
End Sub
Private Sub xGroup_Change()
If xGroup.MatchedWithList Or Trim(xGroup.BoundText) = "" Then
    myload
    CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub
    Private Sub xGroup_LostFocus()
    If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
    End Sub
