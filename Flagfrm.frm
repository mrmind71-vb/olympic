VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FlagFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   5400
   Begin VB.CommandButton Command2 
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
      Height          =   420
      Left            =   135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   3645
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   5190
      _cx             =   9155
      _cy             =   6429
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
      Rows            =   50
      Cols            =   2
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
      Editable        =   2
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
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   3690
      Width           =   5190
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "≈”„ :"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   555
      End
   End
End
Attribute VB_Name = "FlagFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Public nMin As Long
Dim nZeros As Integer
Dim GRDTABLE As New adodb.Recordset
Dim con As New adodb.Connection
Dim CTABLE As String, CGROUP As String
Private Sub myLoad()
Dim cFilter As String
GRDTABLE.Requery
If Trim(xname.Text) <> "" Then cFilter = MyParnAnd(xname.Text, aPublic(2))
GRDTABLE.Filter = cFilter
i = 0
grid1.Rows = 1
Do Until GRDTABLE.EOF
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = GRDTABLE(aPublic(1)) & ""
    grid1.TextMatrix(grid1.Rows - 1, 1) = GRDTABLE(aPublic(2)) & ""
    GRDTABLE.MoveNext
Loop
grid1.AddItem ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set FlagFrm = Nothing
Err.Clear
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim aInsert(1, 1)

aInsert(0, 0) = aPublic(1)
aInsert(0, 1) = addvalue(grid1.TextMatrix(Row, 0))

aInsert(1, 0) = aPublic(2)
aInsert(1, 1) = addstring(grid1.TextMatrix(Row, 1))

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then
    Dim nValue As Long
    nValue = Newflag(aPublic(0), aPublic(1))
    If nValue < nMin Then nValue = nMin + 1
    grid1.TextMatrix(Row, 0) = nValue
    aInsert(0, 1) = grid1.TextMatrix(Row, 0)
    con.Execute CreateInsert(aInsert, aPublic(0))
Else
    con.Execute CreateUpdate(aInsert, aPublic(0), " WHERE " & aPublic(1) & " = " & grid1.TextMatrix(Row, 0))
End If
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myLoad
End Sub
Private Sub grid1_EnterCell()
If grid1.Row > grid1.Rows - 1 Or Not bedit Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbCritical, "Õ–›") = vbYes Then
            If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then
                con.BeginTrans
                con.Execute "Delete From " & aPublic(0) & " where code =  " & grid1.TextMatrix(grid1.Row, 0)
                con.CommitTrans
            End If
            grid1.RemoveItem grid1.Row
            grid1_EnterCell
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myLoad
End Sub
Private Sub Form_Load()
openCon con
cString = "SELECT  *  from " & aPublic(0) & " order by code"
GRDTABLE.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With grid1
.ColWidth(0) = 1200
.ColWidth(1) = grid1.Width - 1600
.TextMatrix(0, 0) = aPublic(3)
.TextMatrix(0, 1) = aPublic(4)
Me.Caption = aPublic(5)
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
myLoad
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myLoad
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "description required"
        Cancel = True
    End If
End If
End Sub
Private Sub xname_Change()
myLoad
End Sub
