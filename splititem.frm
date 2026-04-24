VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form splitItemFrm 
   Caption         =   " ﬁ”Ì„ «·«’‰«›"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4050
      Width           =   6630
      Begin VB.CommandButton cmdexit 
         Caption         =   "Œ—ÊÃ"
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000005&
         Caption         =   "Õ›Ÿ"
         Height          =   465
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1365
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grid1 
      Height          =   3660
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   6630
      _cx             =   11695
      _cy             =   6456
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
      Rows            =   1
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
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "≈Ã„«·Ì «·ﬁÌ„… :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label xTotalQuant 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   1320
   End
End
Attribute VB_Name = "splitItemFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sGroup As String, nType As Integer, sDesca As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If myreplace Then
    With Purchasefrm.grid1
    .TextMatrix(.Row, 2) = grid1.TextMatrix(1, 1)
    .TextMatrix(.Row, 3) = grid1.TextMatrix(1, 2)
    For i = 2 To grid1.Rows - 2
         nRow = .Row + i - 1
        .AddItem "", nRow
        .TextMatrix(nRow, 1) = grid1.TextMatrix(i, 0)
        .TextMatrix(nRow, 2) = grid1.TextMatrix(i, 1)
        .TextMatrix(nRow, 3) = grid1.TextMatrix(i, 2)
        .TextMatrix(nRow, 4) = grid1.TextMatrix(i, 3)
    Next
    End With
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Fixgrd
myLoad
End Sub
Private Sub myLoad()
With Purchasefrm.grid1
xTotalQuant.Caption = .TextMatrix(.Row, 3)
xTotal.Caption = .TextMatrix(.Row, 6)

grid1.AddItem ""
sDesca = .TextMatrix(.Row, 2)

grid1.TextMatrix(1, 0) = .TextMatrix(.Row, 1)
grid1.TextMatrix(1, 1) = .TextMatrix(.Row, 2)
grid1.TextMatrix(1, 2) = .TextMatrix(.Row, 3)
grid1.TextMatrix(1, 3) = .TextMatrix(.Row, 4)
grid1.TextMatrix(1, 4) = Val(.TextMatrix(.Row, 4)) * Val(.TextMatrix(.Row, 3))
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(.Row, 2)
grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(.Row, 4)
MakeGry
End With
End Sub
Private Function myreplace() As Boolean

nLastitem = Val(GetDesca("Select max(file1_10.item) from file1_10")) + 1
sGroup = GetDesca("select [group] from file1_10 where item = " & MyParn(grid1.TextMatrix(1, 0)))
sSupler = GetDesca("select supler from file1_10 where item = " & MyParn(grid1.TextMatrix(1, 0)))
nType = Val(GetDesca("select Type from file1_10 where item = " & MyParn(grid1.TextMatrix(1, 0))))

On Error Resume Next
con.BeginTrans
con.Execute "update file1_10 set file1_10.desca = " & addstring(grid1.TextMatrix(1, 1)) & " where item = " & MyParn(grid1.TextMatrix(1, 0))
For i = 2 To grid1.Rows - 2
    con.Execute "insert into file1_10(item,desca,[group],supler,[type])" & _
                "values(" & _
                addstring(RetZero(nLastitem, 6)) & "," & _
                addstring(grid1.TextMatrix(i, 1)) & "," & _
                addstring(sGroup) & "," & _
                addstring(sSupler) & "," & _
                nType & _
                ")"
    If Err.Number <> 0 Then Exit For
    grid1.TextMatrix(i, 0) = RetZero(nLastitem, 6)
    nLastitem = nLastitem + 1
Next
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub Fixgrd()
grid1.Cols = 5
grid1.TextMatrix(0, 0) = "ﬂÊœ «·’‰›"
grid1.TextMatrix(0, 1) = "≈”„ «·’‰›"
grid1.TextMatrix(0, 2) = "ﬂ„Ì… «·’‰›"
grid1.TextMatrix(0, 3) = "”⁄— «·’‰›"
grid1.TextMatrix(0, 4) = "«·«Ã„«·Ì"
grid1.ColWidth(0) = 1000
grid1.ColWidth(1) = 3000
grid1.ColWidth(2) = 1000
grid1.ColWidth(3) = 1000
grid1.ColWidth(4) = 1000

grid1.ColHidden(0) = True
For i = 1 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
calcRow Row
End Sub
Private Sub grid1_EnterCell()
'If grid1.Row = 1 Then
'   grid1.Editable = IIf(grid1.Col = 2, flexEDKbdMouse, flexEDNone)
'Else
   grid1.Editable = IIf(grid1.Col = 1 Or grid1.Col = 2, flexEDKbdMouse, flexEDNone)
'End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then
    grid1.AddItem "", grid1.Row
    grid1.TextMatrix(grid1.Row, 1) = sDesca
    grid1.TextMatrix(grid1.Row, 3) = grid1.TextMatrix(1, 3)
End If
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row > 1 Then grid1.RemoveItem grid1.Row
MakeGry
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Row + 1, 1) = sDesca
    grid1.TextMatrix(grid1.Row + 1, 3) = grid1.TextMatrix(1, 3)
    MakeGry
End If
End Sub
Private Function MYVALID() As Boolean
Dim nTotal As Double
nTotal = CalcTotals
If nTotal <> Val(xTotalQuant.Caption) Then
    MsgBox "«·ﬂ„Ì… ·«  ’·Õ"
    Exit Function
End If
With grid1
For i = 2 To grid1.Rows - 2
    If grid1.TextMatrix(i, 1) = "" Then
        MsgBox "«·’‰› ›Ï «·”ÿ— " & i & " ·« ÌÊÃœ ·Â Ê’› "
        Exit Function
    End If
    If Val(grid1.TextMatrix(i, 2)) = 0 Then
        MsgBox "«·’‰› ›Ï «·”ÿ— " & i & " ·«  ÊÃœ ·Â ﬂ„Ì…"
        Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Function CalcTotals() As Double
For i = 1 To grid1.Rows - 2
    CalcTotals = Val(grid1.TextMatrix(i, 2)) + CalcTotals
Next
End Function
Private Sub calcRow(nRow)
grid1.TextMatrix(nRow, 4) = Val(grid1.TextMatrix(nRow, 3)) * Val(grid1.TextMatrix(nRow, 2))
End Sub
Private Sub MakeGry()
grid1.Cell(flexcpBackColor, 1, 0, grid1.Rows - 2, grid1.Cols - 1) = &H80000005
grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &H8000000F
End Sub
