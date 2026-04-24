VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form cost_items 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Break Down"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   13665
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -135
      Top             =   1575
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   3615
      Left            =   135
      TabIndex        =   7
      Top             =   45
      Width           =   13470
      _cx             =   23760
      _cy             =   6376
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorSel    =   8454143
      ForeColorSel    =   -2147483630
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   600
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
      RightToLeft     =   0   'False
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton Cmd_Last 
      Caption         =   ">>"
      Height          =   330
      Left            =   6885
      TabIndex        =   6
      Top             =   4230
      Width           =   465
   End
   Begin VB.CommandButton Cmd_First 
      Caption         =   "<<"
      Height          =   330
      Left            =   7380
      TabIndex        =   5
      Top             =   4230
      Width           =   465
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8145
      TabIndex        =   2
      Top             =   4185
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8145
      TabIndex        =   1
      Top             =   3735
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   135
      TabIndex        =   3
      Top             =   3645
      Width           =   6675
      Begin VB.TextBox xDesce 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1530
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   5010
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descritiption :"
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
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
         Name            =   "MS Sans Serif"
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
Attribute VB_Name = "cost_items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As New ADODB.Recordset, cList As String
Public bEdit As Boolean
Private Sub cmd_quit_Click()
Unload Me
Set cost_items = Nothing
End Sub
Private Sub Cmd_Search_Click()
myload
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Command2_Click()
myload
xDesce.Text = ""
End Sub
Private Sub Form_Activate()
On Error Resume Next
xDesce.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TypeOf ActiveControl Is TextBox Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
cList = StrList("select code,desca from cost_center where code < 6 order by code")
Set Grid1.DataSource = data1
data1.ConnectionString = con.ConnectionString

data2.ConnectionString = con.ConnectionString
data2.RecordSource = "select * from costitems"

myload
If Not bEdit Then Grid1.Editable = flexEDNone
End Sub
Private Sub myload()
Dim cString
cString = "SELECT CODE,DESCA FROM COSTITEMS"
If Trim(xDesce.Text) <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesce.Text, "DESCA")
data1.RecordSource = cString & " Order by Code"
data1.Refresh
Fixgrd
Grid1.additem ""
End Sub
Private Sub Fixgrd()
Grid1.FormatString = "Code|Description"
Grid1.ColWidth(0) = 1000
Grid1.ColWidth(1) = 7000
Grid1.ColWidth(2) = 1000
Grid1.ColHidden(0) = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set invitemsBreakfrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim aInsert(3, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = Val(Grid1.TextMatrix(Row, 0))

aInsert(1, 0) = "DescE"
aInsert(1, 1) = addString(Grid1.TextMatrix(Row, 1))

aInsert(2, 0) = "[VALUE]"
aInsert(2, 1) = Val(Grid1.TextMatrix(Row, 2))

aInsert(3, 0) = "CATEGORY"
aInsert(3, 1) = addValue(Grid1.TextMatrix(Row, 3))

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(Grid1.TextMatrix(Row, 0)) Then
    Grid1.TextMatrix(Row, 2) = Newflag("INV_ITEM_BREAK", "code")
    aInsert(0, 1) = addString(Grid1.TextMatrix(Row, 0))
    con.Execute CreateInsert(aInsert, "INV_ITEM_BREAK")
Else
    con.Execute CreateUpdate(aInsert, "INV_ITEM_BREAK", " WHERE CODE = " & Grid1.TextMatrix(Row, 0))
End If
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'Grid1.EditMaxLength = IIf(Col = 3, 1, 0)
End Sub
Private Sub Grid1_DblClick()
If bPop = True Then
    oForm.myproc2
End If
'oForm.myproc2
End Sub
Private Sub Grid1_EnterCell()
If (Not bEdit) Or Check1.Value = 0 Then Grid1.Editable = flexEDNone Else Grid1.Editable = flexEDKbdMouse
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not bEdit Then Exit Sub
On Error GoTo myerror
If KeyCode = 46 And Grid1.Row <> Grid1.Rows - 1 And Grid1.Row <> 0 Then
    cString = "SELECT Cust_Inv_Item_Break.Item_Break_Code " & _
              " FROM Cust_Inv_Item_Break INNER JOIN INV_ITEM_BREAK ON Cust_Inv_Item_Break.Item_Break_Code = INV_ITEM_BREAK.CODE" & _
              " where cust_inv_item_break.Item_Break_Code  = " & Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)
    If GetDesca(cString) <> "" Then
        MsgBox "Item Aleary found in invoices "
        Exit Sub
    End If
    If MsgBox("delete Current ! are you sure ?", vbOKCancel + vbDefaultButton2) = vbOK Then
        If Trim(Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)) <> "" Then
            CardTable.Seek Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1), adSeekFirstEQ
            If Not CardTable.EOF Then CardTable.Delete
        End If
        Grid1.RemoveItem Grid1.Row
    End If
End If
Exit Sub
myerror:
    MsgBox Err.number & vbCrLf & Err.Description
    Err.Clear
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Row = Grid1.Rows - 1 Then
    Grid1.additem ""
    Grid1.TextMatrix(Row, 1) = "0"
End If
End Sub
Private Function validRow(Row, Col) As Boolean
If Trim(Grid1.TextMatrix(Row, 0)) = "" Then
    MsgBox "Description Required"
    Exit Function
End If
validRow = True
End Function
Private Function GetLastNumber() As Integer
If CardTable.EOF And CardTable.BOF Then Exit Function
CardTable.MoveLast
GetLastNumber = CardTable!code
End Function
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(Grid1.EditText) = "" Then
        MsgBox "Description Required"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Not IsNumeric(Grid1.EditText) Then
        MsgBox "Numeric Value Required "
        Cancel = True
    End If
ElseIf Col = 3 Then
'    If Not (Trim(Grid1.EditText) = "" Or Trim(UCase(Grid1.EditText)) = "E" Or Trim(UCase(Grid1.EditText)) = "M" Or Trim(UCase(Grid1.EditText)) = "S" Or Trim(UCase(Grid1.EditText)) = "O") Then
'        Cancel = True
'    Else
'        Grid1.EditText = Trim(UCase(Grid1.EditText))
'    End If
End If
End Sub
Private Sub myDefine()
xDesce.Text = ""
End Sub
Private Sub Cmd_Last_Click()
Grid1.SetFocus
Grid1.Row = Grid1.Rows - 1
Grid1.ShowCell Grid1.Rows - 1, 0
End Sub
Private Sub Cmd_First_Click()
Grid1.SetFocus
Grid1.Row = 1
Grid1.ShowCell 1, 0
End Sub
Private Sub xCategory_Change()
myload
End Sub
Private Sub xcategory_LostFocus()
If Trim(xCategory.Text) <> "" Then
    data2.Recordset.Find "Desca LIKE " & MyParn(xCategory.Text & "%"), , adSearchForward, adBookmarkFirst
    If Not data2.Recordset.EOF Then xCategory.BoundText = data2.Recordset!code & ""
End If
End Sub

Private Sub xDesce_Change()
myload
End Sub
