VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form clientsGrdfrm 
   Caption         =   "»Ì«‰«  «·⁄„·«¡ «·‰ﬁœÌ"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17160
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   17160
   Begin VB.CommandButton cmdExit 
      Height          =   600
      Left            =   135
      Picture         =   "clientgrd.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6525
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   3690
      TabIndex        =   5
      Top             =   6390
      Width           =   13335
      Begin VB.TextBox xAddress 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   4065
      End
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
         Height          =   375
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   3435
      End
      Begin VB.TextBox xCode 
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
         Height          =   375
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   3435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄‰Ê«‰ :"
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
         TabIndex        =   10
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·⁄„Ì· :"
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
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂÊœ :"
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
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1620
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   3780
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
      Height          =   6315
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   16890
      _cx             =   29792
      _cy             =   11139
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      Caption         =   "⁄œœ «·”Ã·« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6345
      Width           =   2040
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1860
      End
   End
End
Attribute VB_Name = "clientsGrdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, SCODE As String
Public myform
Dim CardTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub myload()
Dim cString As String
On Error GoTo myerror
cString = "SELECT CODE as [«·ﬂÊœ], FILE3_10.DESCA as [«·«”„],[PHONE] as [«· ·Ì›Ê‰],[PHONE2] as [«·„Ê»Ì·],ADDRESS AS [«·⁄‰Ê«‰],F_DATE AS [ «—ÌŒ «Ê·],F_BAL1 AS [«·—’Ìœ]" & _
          " FROM FILE3_10"
If Trim(xDesca.Text) <> "" Then
    cString = cString & turn(cString)
    cString = cString & "("
    cString = cString & "(" & MyParnAnd(xDesca.Text, "FILE3_10.desca") & ")"
    cString = cString & " or "
    cString = cString & "(" & MyParnAnd(xDesca.Text, "FILE3_10.phone") & ")"
    cString = cString & ")"
End If

If Trim(xAddress.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xAddress.Text, "FILE3_10.ADDRESS")
End If

If xCode.Text <> "" Then
    cString = cString & turn(cString) & "code = " & MyParn(RetZero(xCode.Text))
End If
cString = cString & " order by FILE3_10.CODE"
Set DATA11.Recordset = myRecordSet(cString, con)

MyAddItem
lblTotal.Caption = IIf(grid1.Rows < 3, "", grid1.Rows - 2)
lblTotal.Caption = lblTotal.Caption & turn(lblTotal.Caption, " ”Ã·")
Fixgrd
grid1.Row = grid1.Rows - 1
grid1.Col = 0
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub Form_Activate()
If SCODE <> "" Then
    grid1.SetFocus
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'SaveText Me, , Array(xGroup.Name)
Set clientgrdfrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim nCode As Integer
If Not validRow(Row) Then Exit Sub
If grid1.Row = grid1.Rows - 1 Then MyAddItem
If Not myreplace(Row) Then myload
End Sub
Private Sub grid1_DblClick()
If (Not IsEmpty(myform)) And grid1.Col = 0 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    myform.SCODE = grid1.TextMatrix(grid1.Row, 0)
    Unload Me
End If
End Sub
Private Sub grid1_EnterCell()
If (Not bedit) Or (grid1.Col = 0) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 And Trim(grid1.TextMatrix(grid1.Row, 0)) <> RetZero("0") Then
    If MsgBox("Õ–› «·⁄„Ì· !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "delete from FILE3_10 where code = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        grid1.Select grid1.Row, 0
    End If
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
con.RollbackTrans
myload
End Sub
Private Sub Form_Load()
bedit = True
openCon con

Set grid1.DataSource = DATA11
data1.ConnectionString = strCon
With grid1
Dim nRow As Long
myload
If SCODE <> "" Then
    nRow = grid1.FindRow(SCODE, , 0)
    If nRow = -1 Then nRow = grid1.Rows - 1
Else
    nRow = grid1.Rows - 1
End If
grid1.Select nRow, 1
grid1.ShowCell nRow, 1
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "≈”„ «·⁄„Ì· „ÿ·Ê»"
        Cancel = True
    End If
End If
Exit Sub
myerror:
On Error Resume Next
If Err.Number <> 0 Then MsgBox Err.Description
CardTable.CancelUpdate
con.RollbackTrans
myload
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 900
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 7000
.WordWrap = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub

Private Sub xAddress_Change()
myload
End Sub
Private Sub xAddress_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub
Private Sub xCode_Change()
myload
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub

Private Sub xDesca_Change()
myload
End Sub
Private Sub xDesca_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, 0) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, 0) = "" Then
    .RemoveItem .Row
    'CalcTotals
End If
End With
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub MyAddItem()
With grid1
    .AddItem ""
End With
End Sub
Private Function myreplace(Row As Long) As Boolean
Dim aInsert(3, 1)

aInsert(0, 0) = "code"
aInsert(0, 1) = addstring(grid1.TextMatrix(Row, 0))

aInsert(1, 0) = "descA"
aInsert(1, 1) = addstring(grid1.TextMatrix(Row, 1))

aInsert(2, 0) = "[PHONE]"
aInsert(2, 1) = addstring(grid1.TextMatrix(Row, 2))

aInsert(3, 0) = "[ADDRESS]"
aInsert(3, 1) = addstring(grid1.TextMatrix(Row, 3))


On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(Row, 0) = "" Then
    grid1.TextMatrix(Row, 0) = RetZero(Newflag("FILE3_10", "CODE", con))
    aInsert(0, 1) = addstring(grid1.TextMatrix(Row, 0))
    con.Execute CreateInsert(aInsert, "FILE3_10")
    grid1.TextMatrix(Row, 0) = grid1.TextMatrix(grid1.Row, 0)
Else
    con.Execute CreateUpdate(aInsert, "FILE3_10", " WHERE FILE3_10.CODE = " & MyParn(grid1.TextMatrix(Row, 0)))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    grid1.Select Row, Col + 1
ElseIf Row < grid1.Rows - 1 Then
    If Not IsEmpty(myform) Then
        myform.SCODE = grid1.TextMatrix(Row, 0)
        Unload Me
    Else
        grid1.Row = Row + 1
        grid1.Select Row + 1, 1
        grid1.ShowCell Row + 1, 1
    End If
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub xPhone1_Change()
myload
End Sub
Private Sub xPhone1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 1
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xPhone1_GotFocus()
myGotFocus xPhone1
End Sub
Private Sub xPhone1_LostFocus()
myLostFocus xPhone1
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
