VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form impchargefrm 
   Caption         =   "„’«—Ì› «· ﬂ·›… «·«” Ì—«œÌ…"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand CMD_CODE 
      Height          =   315
      Left            =   75
      TabIndex        =   8
      Top             =   2775
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "√ﬂÊ«œ „’«—Ì› ≈” Ì—«œÌ…"
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3210
      Width           =   1950
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3210
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8250
      Begin VB.Label xcode 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   8070
      End
      Begin VB.Label xDoc_No 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   8115
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   1785
      Left            =   45
      TabIndex        =   3
      Top             =   990
      Width           =   8250
      _cx             =   14552
      _cy             =   3149
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
      BackColorBkg    =   -2147483633
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
      TabBehavior     =   1
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
      Height          =   240
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3255
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "≈Ã„«·Ì «·„’«—Ì› :"
      Height          =   240
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3255
      Width           =   1590
   End
End
Attribute VB_Name = "impchargefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Dim grdTable As New ADODB.Recordset

Private Sub CMD_CODE_Click()
ReDim aPublic(5)
aPublic(0) = "FILE7_60CH_CODE"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„’—Ê›"
aPublic(4) = "»Ì«‰ "
aPublic(5) = "√ﬂÊ«œ „’«—Ì› ≈” Ì—«œÌ…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub cmdSave_Click()
Dim aGrid(4, 1)
If MsgBox("√÷«›… «·„’—Ê› «·Ì «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton1) <> vbOK Then Exit Sub
On Error GoTo myerror
CON.BeginTrans
With Grid1
    For i = 1 To .Rows - 2
        aGrid(0, 0) = "Doc_no": aGrid(0, 1) = addstring(impcostfrm.xDoc_No.Text)
        aGrid(1, 0) = "Charge": aGrid(1, 1) = addstring(Grid1.TextMatrix(i, 0))
        aGrid(2, 0) = "desca": aGrid(2, 1) = addstring(Grid1.TextMatrix(i, 1))
        aGrid(3, 0) = "[Value]": aGrid(3, 1) = Val(Grid1.TextMatrix(i, 2))
        aGrid(4, 0) = "row": aGrid(4, 1) = i
        If Grid1.TextMatrix(i, Grid1.Cols - 1) = "" Then
            CON.Execute CreateInsert(aGrid, "file7_60CH")
        Else
            CON.Execute CreateUpdate(aGrid, "FILE7_60CH", " WHERE ID = " & Grid1.TextMatrix(i, Grid1.Cols - 1), Array(-1))
        End If
    Next
End With
CON.Execute "update file7_60h set file7_60h.charge = " & Val(xTotal.Caption) & " where doc_no = " & MyParn(impcostfrm.xDoc_No.Text)
CON.CommitTrans
impcostfrm.xCharge.Caption = Val(xTotal.Caption)
Unload Me
Exit Sub
myerror:
    CON.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
OpenCon CON
With impcostfrm
xDoc_No.Caption = "„” ‰œ —ﬁ„ " & .xDoc_No.Text & Space(50) & " » «—ÌŒ : " & .xDate.Text
xcode.Caption = "«·„Ê—œ : " & .xCodeDesca
grdTable.Open "Select * From FILE7_60CH where Doc_no = " & MyParn(.xDoc_No.Text) & " order by [Row]", CON, adOpenStatic, adLockReadOnly, adCmdText
End With
With Grid1
    .Cols = 3
    .Rows = 1
    .Editable = flexEDKbd
    .FormatString = "«·„’—Ê›|" & "«·»Ì«‰|" & "«·ﬁÌ„…|"
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1000
    .ColHidden(Grid1.Cols - 1) = True
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = MakeString
Do Until grdTable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = grdTable!CHARGE & ""
    .TextMatrix(.Rows - 1, 1) = grdTable!DESCA & ""
    .TextMatrix(.Rows - 1, 2) = grdTable!Value & ""
    .TextMatrix(.Rows - 1, 3) = grdTable!ID & ""
     grdTable.MoveNext
Loop
.AddItem ""
CalcTotals
End With
End Sub
Private Function MakeString()
Dim loctable As New ADODB.Recordset
loctable.Open "File7_60CH_CODE", CON, adOpenStatic, adLockReadOnly, adCmdTable
MakeString = "#  " & ";       "
Do Until loctable.EOF
    MakeString = MakeString & "|#" & loctable!CODE & ";" & loctable!DESCA
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
grdTable.Close
Set grdTable = Nothing
closeCon CON
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
CalcTotals
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And Grid1.Row <> Grid1.Rows - 1 Then Grid1.AddItem "", Grid1.Row
If KeyCode = 46 And Grid1.Row <> Grid1.Rows - 1 Then
    If MsgBox("Õ–› »‰œ «· ﬂ·›…!‰⁄„ «„ ·«", vbOKCancel + vbDefaultButton2) = vbOK Then
        If Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1) <> "" Then
            On Error GoTo myerror
            CON.BeginTrans
            CON.Execute "DELETE  FROM FILE7_60CH WHERE ID = " & Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)
            CON.CommitTrans
        End If
        Grid1.RemoveItem Grid1.Row
        CalcTotals
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
CON.RollbackTrans
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Row = Grid1.Rows - 1 Then Grid1.AddItem ""
End Sub

Private Function MYVALID()
With Grid2
For i = 1 To Grid2.Rows - 2
    If Trim(.TextMatrix(i, 0)) = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„’—›"
        Exit Function
    End If
    If Val(.TextMatrix(i, 2)) = 0 Then
        MsgBox "„’—Ê› »œÊ‰ ﬁÌ„…"
        Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Function CalcTotals()
Dim nTotal As Double
With Grid1
For i = 1 To Grid1.Rows - 2
    nTotal = nTotal + Val(Grid1.TextMatrix(i, 2))
Next
xTotal.Caption = Format(nTotal, "Fixed")
End With
End Function

