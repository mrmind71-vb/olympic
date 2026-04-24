VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Vs_Bank 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»‰ﬂ"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Cmd_Inv 
      BackColor       =   &H00E0E0E0&
      Caption         =   " ”ÃÌ· »‰Êœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   945
      Width           =   1365
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E0E0E0&
      Caption         =   " —«Ã⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1537
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   945
      Width           =   1365
   End
   Begin VB.CommandButton CmdDelInv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ–› «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2925
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton CmdAddItem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "≈÷«›… »‰Êœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8025
      MaskColor       =   &H00E0E0E0&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton cmdNewinv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "„” ‰œ ÃœÌœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
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
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ›Ÿ «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   945
      Width           =   1365
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1290
   End
   Begin VB.TextBox xDoc_No 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6675
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   555
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   4890
      Left            =   75
      TabIndex        =   3
      Top             =   1440
      Width           =   9285
      _cx             =   16378
      _cy             =   8625
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      BackColorFixed  =   12615808
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
      BackColorAlternate=   14737632
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
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Vs_Bank.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   3
      OutlineCol      =   1
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
      TextStyleFixed  =   1
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin MSDBCtls.DBCombo xBank 
      Bindings        =   "Vs_Bank.frx":0088
      DataSource      =   "Data1"
      Height          =   315
      Left            =   5385
      TabIndex        =   14
      Top             =   960
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»‰ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·√Ã„«·Ï"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6480
      Width           =   705
   End
   Begin VB.Label LblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6480
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      FillColor       =   &H00C0C0C0&
      FillStyle       =   6  'Cross
      Height          =   405
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   555
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   795
   End
End
Attribute VB_Name = "Vs_Bank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset
Dim COUNTINVTOTAL As Double
Dim BoxTable As Recordset
Dim formMode, cStrBox As String, cType As String
Dim myFileName As String
Const NewInvMode = 4, applyMode = 5
Sub DocValid()
If xDoc_No.Text = "" Then Exit Sub
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.FindFirst " doc_no = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    Exit Sub
Else
    ApplyProc
End If
End Sub
Sub EmptyProc()
formMode = EmptyMode
ItemInv.Rows = 1
End Sub
Sub AddProc()
formMode = addmode
ItemInv.AddItem ""
End Sub
Sub Fillgrd()
COUNTINVTOTAL = 0
ItemInv.Rows = 1
I = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow
LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(I, 0) = TurnValue(DocTable.BOX, Null, "")
    .TextMatrix(I, 1) = TurnValue(DocTable!Type, Null, "")
    .TextMatrix(I, 2) = TurnValue(DocTable.DESCA, Null, "")
    .TextMatrix(I, 3) = TurnValue(Format(DocTable!Value, "###0.00"), Null, "")
     COUNTINVTOTAL = COUNTINVTOTAL + DocTable.Value
    LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
    DocTable.MoveNext
    If DocTable.EOF Then Exit Sub
    If DocTable.DOC_NO <> xDoc_No.Text Then Exit Sub
    I = I + 1
Loop
End With
End Sub
Sub ApplyProc()
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    EmptyProc
    xDoc_No.Enabled = True
Else
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    xBank.BoundText = DocTable.bank
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
'    ItemInv.EditText = GrdText(Search.Grid1, 0)
'    ItemInv.TextMatrix(ItemInv.row, 1) = GrdText(Search.Grid1, 0)
'    ItemInv.TextMatrix(ItemInv.row, 2) = GrdText(Search.Grid1, 1)
'    ChargTable.FindFirst " CODE = " & MyParn(ItemInv.EditText)
'    If Not ChargTable.NoMatch Then
'        ItemInv.TextMatrix(ItemInv.row, 2) = ChargTable.DESCA
'    End If
Else
    ActiveControl.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·„” ‰œ"
    MYVALID = False
End If
If xDate.Text = "" Or Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
If xBank.BoundText = "" Then
    MsgBox " ”ÃÌ· «·»‰ﬂ"
    MYVALID = False
End If
End Function
Sub Undoinv()
Select Case formMode
    Case addmode
        InvGrid.Rows = InvGrid.Rows - 1
        dispProc
    Case Editmode
        dispProc
    Case EmptyMode
End Select
End Sub
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 2
End Sub
Private Sub cmdAdditem_Click()
    With ItemInv
    .Rows = .Rows + 1
    For I = ItemInv.Rows - 1 To 2 Step -1
        .TextMatrix(I, 0) = .TextMatrix(I - 1, 0)
        .TextMatrix(I, 1) = .TextMatrix(I - 1, 1)
        .TextMatrix(I, 2) = .TextMatrix(I - 1, 2)
        .TextMatrix(I, 3) = .TextMatrix(I - 1, 3)
    Next I
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = ""
    .TextMatrix(1, 2) = ""
    .TextMatrix(1, 3) = ""
    End With
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ  »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        xDate.Text = ""
        xBank.BoundText = ""
        Fillgrd
        xDoc_No.Enabled = True
        ItemInv.Enabled = False
    End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
If Not MyReplace Then Exit Sub
ItemInv.Rows = 1
ItemInv.AddItem ""
xDate.Text = Date
xBank.BoundText = Date
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Fillgrd
End Sub
Private Sub CmdUndo_Click()
    DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not DocTable.NoMatch Then
        xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
        Fillgrd
    End If
End Sub
Private Sub Form_Load()
Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE5_30  ORDER BY DOC_NO ", dbOpenDynaset)
Set BoxTable = mydb.OpenRecordset("SELECT * FROM file0_50 ORDER BY CODE ", dbOpenDynaset)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "FILE5_10"
xBank.BoundColumn = "ID"
xBank.ListField = "DESCA"
cType = "#2;≈Ìœ«⁄ »‰ﬂÏ|#3;”Õ» »‰ﬂÏ|#4;„’«—Ì› »‰ﬂÌ…"
cStrBox = StrBox
xDate.Text = Format(Date, "dd-mm-yyyy")
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 3
    .Rows = 1
    .Editable = flexEDKbdMouse
    .FormatString = "Œ“‰…|" & "‰Ê⁄ «·Õ—ﬂ…|" & "»Ì«‰|" & "«·≈Ã„«·‹‹‹Ï"
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1500
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTDouble
    
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColComboList(0) = cStrBox
    .ColComboList(1) = cType
    .MergeCells = 2
End With
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("Õ–› ”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        ItemInv.RemoveItem ItemInv.Row
    End If
End If
End Sub
Private Sub ItemInv_SelChange()
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col + 1 = ItemInv.Cols Then
        ItemInv.Rows = ItemInv.Rows + 1
    End If
End Sub
Private Sub ItemInv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
    Select Case ItemInv.Col
        Case 1
            If ItemInv.EditText = "" Then
                Cancel = True
            End If
    End Select
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    Generalarray(2) = "SELECT FILE5_30.DOC_NO, FILE5_30.DATE, FILE5_10.DESCA " & _
                      " FROM FILE5_10 RIGHT JOIN FILE5_30 ON FILE5_10.ID = FILE5_30.CODE "
    Generalarray(3) = " Where DescA Like '*cFilter*'"
    Generalarray(4) = " ORDER BY DATE "
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_LostFocus()
DocValid
End Sub
Function myDelete()
    ' Õ–›  «·„” ‰œ
    cString = " DELETE * FROM FILE5_30  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE5_30  ORDER BY DOC_NO ", dbOpenDynaset)
End Function
Function MyReplace()
MyReplace = True
' „—«Ã⁄… «·„” ‰œ
With ItemInv
For I = 1 To .Rows - 1
'    If .TextMatrix(I, 1) = "" AND Then
'        .Select I, 0, I, 3
'        cMess = "·‰ Ì „  ”ÃÌ· Â–« «·»‰œ   "
'        MsgBox cMess
'        MyReplace = False
'    End If
Next I
End With
If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  * FROM FILE5_30  WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    '  ”ÃÌ· «·„” ‰œ
    With ItemInv
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(I, 3)) <> 0 And .TextMatrix(I, 1) <> "" Then
        DocTable.AddNew
        DocTable.DOC_NO = xDoc_No.Text
        DocTable.bank = xBank.BoundText
        DocTable.[Date] = xDate.Text
        DocTable!Type = .TextMatrix(I, 1)
        DocTable!Value = Val(.TextMatrix(I, 3))
        DocTable.BOX = TurnValue(.TextMatrix(I, 0), "", Null)
        DocTable.DESCA = TurnValue(.TextMatrix(I, 2), "", Null)
        DocTable.Update
        End If
    Next I
    
    ' ≈·€«¡ Õ—ﬂ… «·»‰ﬂ
    cString = " DELETE  * FROM FILE5_11  WHERE DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '2' "
    mydb.Execute cString
    cString = " DELETE  * FROM FILE5_11  WHERE DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '3' "
    mydb.Execute cString
    cString = " DELETE  * FROM FILE5_11  WHERE DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
    mydb.Execute cString
    
    ' Õ—ﬂ… «·»‰ﬂ
    cString = "INSERT INTO FILE5_11( " & _
               "[Date],Doc_Id,[Type],BANK,DescA,DescMove,[IN])" & _
               " Select [Date],Doc_No,'2'," & _
               "BANK,DESCA,'≈Ìœ«⁄ »‰ﬂÏ'," & _
               " [VALUE]" & _
               " From File5_30 " & _
               " WHERE FILE5_30.DOC_NO = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '2' "
    mydb.Execute cString
    
    cString = "INSERT INTO FILE5_11( " & _
               "[Date],Doc_Id,[Type],BANK,DescA,DescMove,[OUT])" & _
               " Select [Date],Doc_No,'3'," & _
               "BANK,DESCA,'”Õ» »‰ﬂÏ'," & _
               " [VALUE]" & _
               " From File5_30 " & _
               " WHERE FILE5_30.DOC_NO = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '3' "
    mydb.Execute cString
    
    cString = "INSERT INTO FILE5_11( " & _
               "[Date],Doc_Id,[Type],BANK,DescA,DescMove,[OUT])" & _
               " Select [Date],Doc_No,'4'," & _
               "BANK,DESCA,'„’«—Ì› »‰ﬂÌ…'," & _
               " [VALUE]" & _
               " From File5_30 " & _
               " WHERE FILE5_30.DOC_NO = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
    mydb.Execute cString
    
    End With
    Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE5_30  ORDER BY DOC_NO ", dbOpenDynaset)
End If
End Function
Private Function StrBox()
If BoxTable.RecordCount > 0 Then
    BoxTable.MoveFirst
    I = 1
    StrBox = "#" & BoxTable!CODE & ";" & BoxTable.DESCA
    BoxTable.MoveNext
    Do While True
        I = I + 1
        If BoxTable.EOF Then Exit Do
        StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable.DESCA
        BoxTable.MoveNext
    Loop
End If
End Function
Private Sub ItemInv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col = 0 And (ItemInv.TextMatrix(ItemInv.Rows - 1, 1) = "") Then
        ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = ItemInv.TextMatrix(ItemInv.Rows - 2, 0)
        ItemInv.TextMatrix(ItemInv.Rows - 1, 1) = ItemInv.TextMatrix(ItemInv.Rows - 2, 1)
    End If
End Sub
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 3 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub

