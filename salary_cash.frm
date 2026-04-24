VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form salary_cashfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   11025
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   10770
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1035
         Width           =   1590
      End
      Begin VB.Label Label13 
         Caption         =   "«·»«ﬁÌ"
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   1590
      End
      Begin VB.Label Label11 
         Caption         =   "«·„”œœ"
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì «·„— »"
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "«·„ÊŸ›"
         Height          =   330
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "«·«”»Ê⁄"
         Height          =   285
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7065
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "«·‘Â—"
         Height          =   330
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8910
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "«·”‰…"
         Height          =   330
         Left            =   9765
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   600
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5565
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "salary_cash.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4995
      UseMaskColor    =   -1  'True
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   -45
      Top             =   -270
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
      Height          =   3435
      Left            =   135
      TabIndex        =   0
      Top             =   1485
      Width           =   10725
      _cx             =   18918
      _cy             =   6059
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Cols            =   9
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
      WordWrap        =   -1  'True
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
      Left            =   675
      Top             =   585
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
End
Attribute VB_Name = "salary_cashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Public sDoc_no  As String
Dim cList As String, cList2 As String, nMode As Integer
Private Sub MyloadGrd()
Dim cString
cString = "SELECT ENG,CHARGE,TOTAL, ID FROM ENG_CHARGE"
cString = cString & turn(cString) & "DOC_NO = " & sDoc_no
If xEng.MatchedWithList Then cString = cString & turn(cString) & "eng = " & xEng.BoundText
data11.RecordSource = cString
data11.Refresh
myAddItem
Fixgrd
CALCTOTALS
End Sub
Private Sub cmdcopy_Click()
If grid1.Row = grid1.Rows - 1 Then Exit Sub
For I = grid1.Row To grid1.Rows - 2
    For nCol = 0 To grid1.Cols - 2
        grid1.TextMatrix(grid1.Rows - 1, nCol) = grid1.TextMatrix(I, nCol)
        grid1_AfterEdit grid1.Rows - 1, 0
    Next
Next
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 And Not bStill Then
    KeyCode = 0
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set arrivefrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then myAddItem

con.BeginTrans
On Error GoTo myerror
Dim aInsert As Variant
aInsert = AddFlag(Empty, "DOC_NO", addvalue(sDoc_no))
aInsert = AddFlag(aInsert, "Eng", addstring(.TextMatrix(Row, 0)))
aInsert = AddFlag(aInsert, "charge", addvalue(.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "Total", Val(.TextMatrix(Row, 2)))
If grid1.TextMatrix(Row, .Cols - 1) = "" Then
   con.Execute addInsert(aInsert, "ENG_CHARGE")
Else
   con.Execute addUpdate(aInsert, "ENG_CHARGE", "ID = " & .TextMatrix(Row, .Cols - 1))
End If
con.CommitTrans
If grid1.TextMatrix(Row, .Cols - 1) = "" Then MyloadGrd Else CALCTOTALS
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
MyloadGrd
End Sub
Private Sub grid1_EnterCell()
If bedit Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 Then
    If MsgBox("Delete are you sure", vbYesNo) = vbYes Then
        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            Dim cString As String
            cString = "Delete from eng_charge"
            cString = cString & turn(cString) & " ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.Execute cString
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        CALCTOTALS
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Form_Load()
bedit = True
Dim cString As String

cString = "SELECT STAFF.CODE, STAFF.DESCA " & _
           "FROM STAFF" & _
           " WHERE STAFF.CODE IN (SELECT Eng_Code FROM Cust_Inv_Engineers WHERE INV_NO = " & sDoc_no & ")" & _
           " OR STAFF.CODE IN (SELECT ADMIN FROM Cust_Invoices_Head WHERE INVNO = " & sDoc_no & ")" & _
           " OR STAFF.CODE IN (SELECT DRIVER FROM Cust_Invoices_Head WHERE INVNO = " & sDoc_no & ")" & _
           " ORDER BY STAFF.DESCA"

cList = StrList(cString)
cList2 = StrList("SELECT code,desca from charges_codes")

DATA1.ConnectionString = con.ConnectionString
DATA1.RecordSource = cString
Set xEng.RowSource = DATA1
xEng.ListField = "CODE"
xEng.BoundColumn = "DESCA"

Set grid1.DataSource = data11
data11.ConnectionString = con.ConnectionString
MyloadGrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
End If
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CALCTOTALS
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, grid1.Cols - 1) = "" Then
    .RemoveItem .Row
    CALCTOTALS
End If
End With
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
If Val(grid1.TextMatrix(Row, 2)) = 0 Then Exit Function
validRow = True
End With
End Function
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 4000
.ColWidth(1) = 4000
.ColWidth(2) = 1200
.TextMatrix(0, 0) = "Engineer"
.TextMatrix(0, 1) = "Charge"
.TextMatrix(0, 2) = "Total"
.ColComboList(0) = cList
.ColComboList(1) = cList2
.RowHeight(0) = 600
.ColHidden(.Cols - 1) = True
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignLeftCenter
Next
End With
End Sub
Private Sub xeng_Change()
If xEng.MatchedWithList Or Trim(xEng.BoundText) = "" Then
    MyloadGrd
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 And Col <> 1 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    grid1.ShowCell Row + 1, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub myAddItem()
grid1.AddItem ""
If grid1.Rows > 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.TextMatrix(grid1.Rows - 2, 0)
End If
End Sub
Private Sub CALCTOTALS()
Dim nTotal As Long
For I = 1 To grid1.Rows - 2
    nTotal = nTotal + Val(grid1.TextMatrix(I, 2))
Next
If nTotal <> 0 Then
    Me.StatusBar1.Panels(1).Text = "Total Charges : " & nTotal
Else
    Me.StatusBar1.Panels(1).Text = ""
End If
End Sub

