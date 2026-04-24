VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addModelFrm 
   Caption         =   "«÷«›… Ê ⁄œÌ· „ÊœÌ·«  «· ÕÊÌ·« "
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   13470
      _cx             =   23760
      _cy             =   8096
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   3
      FixedCols       =   2
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdSave 
         Caption         =   " ”ÃÌ· ﬂ„Ì«  «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1770
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   300
      Top             =   -540
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   5895
      TabIndex        =   4
      Top             =   90
      Width           =   7575
      Begin VB.TextBox xModel 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4860
         TabIndex        =   8
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   4740
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   3285
      Left            =   45
      TabIndex        =   7
      Top             =   2115
      Width           =   13425
      _cx             =   23680
      _cy             =   5794
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   4230
      Top             =   -90
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5355
      Width           =   2535
      Begin VB.CommandButton Cmd_Item 
         Caption         =   " ⁄œÌ· »Ì«‰«  „ÊœÌ·"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   2355
      End
   End
End
Attribute VB_Name = "addModelFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cDoc_no As String, myFlag As Integer
Public strX1 As String, strX2 As String, strX3 As String, strX4 As String, strModel As String, strModelDesca As String
Dim bUpDate As Boolean
Const LoadMode = 0, DefineMode = 1
Private Sub CMD_SHOW_Click()
    If xModel.Text = "" Then Exit Sub
    Load ShowModl
    ShowModl.xModel.Text = xModel.Text
    ShowModl.AllMoveModel
    ShowModl.Show
End Sub
Private Sub cmd_item_Click()
Dim cString As String
If xX1.BoundText <> "" And xX2.BoundText <> "" And xX3.BoundText <> "" And xX4.BoundText <> "" Then
    ITEMS.strX1 = xX1.Text
    ITEMS.strX2 = xX2.Text
    ITEMS.strX3 = xX3.Text
    ITEMS.strX4 = xX4.Text
    ITEMS.strDesca = xdesca.Caption
    ITEMS.strModel = xModel.Caption
    ITEMS.StrCode = Purchasefrm.xCode.Text
    ITEMS.strCodeDesca = Purchasefrm.xCodeDesca.Caption
    ITEMS.Show 1
    
    cString = xX4.BoundText
    Call FilterModelList
    xX4.BoundText = cString
    xX4.Text = cString
    myloadgrd
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
Set addModelFrm = Nothing
End Sub
Private Function myreplace() As Boolean
Dim cString As String, nRow As Integer, cFile As String
On Error GoTo myerror

cFile = IIf(myFlag = 0, "FILE7_20", "FILE7_10")
cString = " DELETE  FROM " & cFile & " WHERE DOC_NO = " & MyParn(cDoc_no) & " and MODEL = " & MyParn(xModel.Caption)

con.BeginTrans
con.Execute cString

With grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(grid1.TextMatrix(nRow, nCol)) <> 0 Then
                aret = aGetDesca("SELECT item,BarCode,cost,cost2 from file1_10 where item = " & MyParn(Grid2.TextMatrix(nRow, nCol)))
                If UBound(aret) > 0 Then
                    cString = "Insert into " & cFile & " (doc_no,Model,item,price,cost2,Quant)" & _
                               "Values(" & _
                               addstring(cDoc_no) & "," & _
                               addstring(xModel.Caption) & "," & _
                               addstring(aret(1)) & "," & _
                               Val(aret(3)) & "," & _
                               Val(aret(4)) & "," & _
                               Val(grid1.TextMatrix(nRow, nCol)) & _
                               ")"
                    con.Execute cString
                End If
            End If
        Next
    Next
End With
con.CommitTrans
myreplace = True
Exit Function
myerror:
    MsgBox Err.Description
    con.RollbackTrans
    Err.Clear
End Function
Private Sub CMD_EXIT_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Purchasefrm.xDoc_No.Tag = DefineMode Then
    If Not Purchasefrm.mySave(False) Then Exit Sub
End If
    
If myreplace Then
    Inform " „ Õ›Ÿ «·»Ì«‰ »‰Ã«Õ"
    Purchasefrm.myloadgrd
    mydefine
End If
On Error Resume Next
xX1.SetFocus
Err.Clear
End Sub
Private Sub Form_Load()
Cmd_Item.Enabled = myFlag = 0
data1.ConnectionString = strCon
data1.RecordSource = "SELECT CODE, DESCA FROM TABLE_X1 ORDER BY CODE"
Set xX1.RowSource = data1
xX1.ListField = "CODE"
xX1.BoundColumn = "DESCA"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT CODE, DESCA FROM TABLE_X2 ORDER BY CODE"
Set xX2.RowSource = DATA2
xX2.ListField = "CODE"
xX2.BoundColumn = "DESCA"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "SELECT CODE, DESCA FROM TABLE_X3 ORDER BY CODE"
Set xX3.RowSource = DATA3
xX3.ListField = "CODE"
xX3.BoundColumn = "DESCA"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT X4 , MODEL FROM FILE1_10H"
Set xX4.RowSource = data4
xX4.ListField = "X4"
xX4.BoundColumn = "MODEL"

If strX1 <> "" Or strX2 <> "" Or strX3 <> "" Or strX4 <> "" Or strModel <> "" Then
    xX1.Text = strX1
    xX2.Text = strX2
    xX3.Text = strX3
    xX4.Text = strX4
    lblX1.Caption = xX1.BoundText
    LBLX2.Caption = xX2.BoundText
    lblX3.Caption = xX3.BoundText
    xModel.Caption = strModel
    xdesca.Caption = strDesca
    
    myloadgrd
Else
    myDefineGrd
End If
End Sub
Private Sub VsModel_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim nTot As Double
    With VsModel
'        If .Col = 2 Then
'            For c = 2 To .Cols - 1
'                .TextMatrix(.Row, c) = .TextMatrix(.Row, 2)
'            Next c
'        End If
        For r = 4 To .Rows - 1
            For c = 2 To .Cols - 1
                nTot = nTot + Val(.TextMatrix(r, c))
            Next c
        Next r
    End With
    xTot.Caption = nTot
End Sub
Private Sub VsModel_EnterCell()
With VsModel
    .Cell(flexcpBackColor, 4, 2, .Rows - 1, .Cols - 1) = &HFFFFFF
    .Cell(flexcpBackColor, .Row, .Col) = &HFFC0C0

    ItemTable.Index = "nModel"
    ItemTable.Seek "=", xModel.Text, .TextMatrix(.Row, 0), .TextMatrix(1, .Col)
    If Not ItemTable.NoMatch Then
        xBalNo.Caption = BalNoItem(ItemTable.Item, Vs_Inv.xstore.Text)
    Else
        xBalNo.Caption = ""
    End If
End With
End Sub
Private Sub VsModel_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = RetNumber(KeyAscii, False)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then SendKeys "{tAB}"
End Sub
Private Sub VsModel_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VsModel
If publicFlaginv = 2 Then
    If Val(.EditText) > Val(xBalNo.Caption) Then
        MsgBox "·« Ì„ﬂ‰  ”ÃÌ· «·„— Ã⁄ - «·—’Ìœ ·« Ì”„Õ »–·ﬂ"
        Cancel = True
    End If
End If
End With

End Sub

Private Sub xX1_Change()
myloadModel
End Sub

Private Sub xX1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And xX1.Text = "" Then Unload Me
End Sub
Private Sub xX1_Validate(Cancel As Boolean)
If Not xX1.MatchedWithList Then
    xX1.BoundText = ""
    lblX1.Caption = ""
Else
    lblX1.Caption = xX1.BoundText
End If
Call FilterModelList
End Sub

Private Sub xX2_Change()
myloadModel
End Sub

Private Sub xX2_LostFocus()
If Not xX2.MatchedWithList Then
    xX2.BoundText = ""
    LBLX2.Caption = ""
Else
    LBLX2.Caption = xX2.BoundText
End If
Call FilterModelList
End Sub

Private Sub xX3_Change()
myloadModel
End Sub

Private Sub xX3_Validate(Cancel As Boolean)
If Not xX3.MatchedWithList Then
    xX3.BoundText = ""
    lblX3.Caption = ""
Else
    lblX3.Caption = xX3.BoundText
End If
Call FilterModelList
End Sub
Private Sub xX4_Change()
myloadModel
End Sub

Private Sub xX4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If xX4.MatchedWithList Then
        myloadgrd
    Else
        cmd_item_Click
    End If
End If
End Sub
Private Sub VsModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 19 Then
    cmdSave_Click
    xX1.SetFocus
End If
End Sub
Private Sub myload()
    ItemTable.Index = "nItem"
    With invTable
    .Index = "nModel"
    .Seek "=", Vs_Inv.xDoc_No.Text, xModel.Text
    If Not .NoMatch Then
        Do While .doc_no = Vs_Inv.xDoc_No.Text And UCase(.Model) = UCase(xModel.Text)
            ItemTable.Seek "=", .Item
            For r = 4 To VsModel.Rows - 2
                If VsModel.TextMatrix(r, 0) = ItemTable.C_COLOR Then Exit For
            Next r
            
            For c = 2 To VsModel.Cols - 2
                If VsModel.TextMatrix(1, c) = ItemTable.Scal Then Exit For
            Next c
            VsModel.TextMatrix(r, c) = Val(VsModel.TextMatrix(r, c)) + .Quant
            .MoveNext
            If .EOF Then Exit Do
        Loop
    End If
    End With
End Sub
Sub ShowModelInv()
    If GetDesca("select Model from file1_10h where Model =  " & MyParn(xModel.Caption)) = "" Then
        If flag = 0 Then
            ITEMS.strX1 = xX1.Text
            ITEMS.strX2 = xX2.Text
            ITEMS.strX3 = xX3.Text
            ITEMS.strX4 = xX4.Text
            ITEMS.strModel = xModel.Caption
            ITEMS.strDesca = xdesca.Caption
            ITEMS.StrCode = Purchasefrm.xCode.Text
            ITEMS.strCodeDesca = Purchasefrm.xCodeDesca.Caption
            ITEMS.Show 1
        End If
    Else
        xdesca.Text = xModel.Text
        myloadgrd
        VsModel.Select 3, 1
        VsModel.ShowCell 3, 1
        VsModel.SetFocus
    End If
End Sub
Private Sub FilterModelList()
Dim cString As String
cString = "SELECT X4 , MODEL FROM FILE1_10H"
If Trim(xX1.BoundText) <> "" Then cString = cString & turn(cString) & "X1 = " & MyParn(xX1.Text)
If Trim(xX2.BoundText) <> "" Then cString = cString & turn(cString) & "X2 = " & MyParn(xX2.Text)
If Trim(xX3.BoundText) <> "" Then cString = cString & turn(cString) & "X3 = " & MyParn(xX3.Text)
data4.RecordSource = cString
data4.Refresh
End Sub
Private Sub myloadgrd()
Dim aret As Variant, cFieldas As String, cField As String

myDefineGrd

aret = retFields
If aret(0) = "" Then Exit Sub

cField = aret(0)
cFieldas = aret(1)

FillItem cFieldas, cField
FixCost cFieldas, cField
Fixgrd
End Sub
Private Sub Fixgrd()
grid1.ColWidth(0) = 500
grid1.ColWidth(1) = 1200
nColWidth = (grid1.Width - 200 - grid1.ColWidth(0) - grid1.ColWidth(1)) / grid1.Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1000 Then nColWidth = 1000
For nCol = 2 To grid1.Cols - 1
    grid1.ColWidth(nCol) = nColWidth
    grid1.ColAlignment(nCol) = flexAlignCenterCenter
Next
End Sub
Private Sub myDefineGrd()
Grid2.Rows = 3
Grid2.Cols = 2

grid1.Rows = 3
grid1.Cols = 2

grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
'grid1.FixedRows = 3
End Sub
Private Sub myloadModel()
If Trim(xX1.Text) <> "" And Trim(xX2.Text) <> "" And Trim(xX3.Text) <> "" And Trim(xX4.Text) <> "" Then
    xModel.Caption = UCase(Trim(xX1.Text)) & "/" & UCase(Trim(xX2.Text)) & "/" & UCase(Trim(xX3.Text)) & "/" & UCase(Trim(xX4.Text))
    xdesca.Caption = lblX1.Caption & Space(1) & LBLX2.Caption & Space(1) & lblX3.Caption & Space(1) & xX4.Text
    Cmd_Item.Enabled = True
    'myloadgrd
Else
    xModel.Caption = ""
    xdesca.Caption = ""
    myDefineGrd
    Cmd_Item.Enabled = False
End If
End Sub
Private Sub FixCost(cFieldas, cField, Optional cFieldAdd As String = "Cost", Optional nRow As Integer = 1)
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal," & cFieldAdd & " from file1_10 WHERE MODEL = " & MyParn(xModel.Caption) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(" & cFieldAdd & ")" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim locTable As New ADODB.Recordset
locTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not locTable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(nRow, nCol) = locTable.Fields(nCol - 2).Value & ""
    Next
End If
locTable.Close
Set locTable = Nothing
End Sub
Private Function retFields()
Dim aret(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(xModel.Caption) & " group by SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!Scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!Scal & "]" & " as " & "[" & FieldTable!Scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!Scal & "]"
    End If
    FieldTable.MoveNext
Loop

aret(0) = cField
aret(1) = cFieldas
retFields = aret
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub FillItem(cFieldas, cField)
Dim GRDTABLE As New ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item,col_color from file1_10 WHERE MODEL = " & MyParn(xModel.Caption) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.col_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GRDTABLE.Fields.Count: Grid2.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
Next

Do Until GRDTABLE.EOF
    Grid2.AddItem ""
    grid1.AddItem ""
    For nCol = 0 To GRDTABLE.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(Grid2.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
        Else
            Grid2.TextMatrix(Grid2.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
            nFoundRow = Purchasefrm.grid1.FindRow(GRDTABLE.Fields(nCol).Value & "", , 0)
            If nFoundRow <> -1 Then
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = Purchasefrm.grid1.TextMatrix(nFoundRow, 6)
             Else
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
             End If
        End If
    Next
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub mydefine()
xX1.Text = ""
xX2.Text = ""
xX3.Text = ""
xX4.Text = ""
lblX1.Caption = ""
LBLX2.Caption = ""
lblX3.Caption = ""
xModel.Caption = ""
xdesca.Caption = ""
myDefineGrd
End Sub


