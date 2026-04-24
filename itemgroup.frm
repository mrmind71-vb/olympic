VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form itemsGroupFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„Ã„Ê⁄«  «·√’‰«›"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   8385
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
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   1530
      TabIndex        =   2
      Top             =   3735
      Width           =   4515
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo XGROUP 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
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
         Left            =   2685
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   2685
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   555
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3690
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   8295
      _cx             =   14631
      _cy             =   6509
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "itemsGroupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Dim con As New adodb.Connection
Dim clist1 As String, cList2 As String
Dim CardTable As New adodb.Recordset
Private Sub myLoad()
cString = "SELECT CODE as [«·ﬂÊœ], DESCA as [«·»Ì«‰],[GROUP] as [«·„Ã„Ê⁄…],SHOW AS [«ŸÂ«—]" & _
          " FROM FILE1_50"
If IsNumeric(xGroup.BoundText) Then
    cString = cString & turnFound(cString) & " [GROUP] = " & xGroup.BoundText
End If

If Trim(xDesca.Text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xDesca.Text, "desca")
End If

cString = cString & " order by CODE"
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
If xGroup.BoundText <> "" Then
   grid1.TextMatrix(grid1.Rows - 1, 2) = xGroup.BoundText
End If
Fixgrd
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set FlagFrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim aInsert(3, 1)

aInsert(0, 0) = "Code"
aInsert(0, 1) = addvalue(grid1.TextMatrix(Row, 0))

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(grid1.TextMatrix(Row, 1))

aInsert(2, 0) = "[Group]"
aInsert(2, 1) = addvalue(grid1.TextMatrix(Row, 2))

aInsert(3, 0) = "[SHOW]"
aInsert(3, 1) = IIf(Val(grid1.TextMatrix(Row, 3)) = 0, "0", "1")

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then
    Dim nValue As Long
    nValue = Newflag("FILE1_50", "CODE")
'    If nValue < 1000 Then nValue = 1001
    grid1.TextMatrix(Row, 0) = nValue
    aInsert(0, 1) = grid1.TextMatrix(Row, 0)
    con.Execute CreateInsert(aInsert, "FILE1_50")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_50", " WHERE FILE1_50.code = " & grid1.TextMatrix(Row, 0))
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
If (grid1.Col = 0) Then
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
                con.Execute "Delete From file1_50 where code =  " & grid1.TextMatrix(grid1.Row, 0)
                con.CommitTrans
            End If
            grid1.RemoveItem grid1.Row
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
data2.ConnectionString = strCon
data2.RecordSource = "FILE1_50G"
Set xGroup.RowSource = data2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set grid1.DataSource = data1
data1.ConnectionString = strCon
With grid1
clist1 = StrList("Select code,desca from file1_50G order by desca")
myLoad
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
If xGroup.BoundText <> "" Then
   grid1.TextMatrix(grid1.Rows - 1, 2) = xGroup.BoundText
End If
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myLoad
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
End If
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "Ê’› «·„Ã„Ê⁄… „ÿ·Ê»"
        Cancel = True
    End If
End If
On Error GoTo myerror
'    If Col = 2 Then
'        If Trim(GetDesca("Select Desca From FILE1_50G where Desca = " & MyParn(grid1.EditText))) = "" Then
'            If MsgBox("«÷«›… „Ã„Ê⁄… ÃœÌœ…", vbYesNo + vbDefaultButton1) = vbYes Then
'                con.BeginTrans
'                On Error GoTo myerror
'                nCode = Newflag("FILE1_50", "code")
'                con.Execute "insert into FILE1_50G(CODE,Desca) " & _
'                "values(" & _
'                addvalue(nCode) & "," & _
'                addstring(grid1.EditText) & _
'                ")"
'                con.CommitTrans
'                grid1.EditText = nCode
'                clist1 = StrList("select * from file1_50G order by desca")
'                grid1.ColComboList(2) = clist1
'                DATA2.Refresh
'            Else
'                Cancel = True
'            End If
'        End If
'    End If
Exit Sub
myerror:
On Error Resume Next
If Err.Number <> 0 Then MsgBox Err.Description
CardTable.CancelUpdate
con.RollbackTrans
myLoad
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.ColComboList(2) = clist1
.ColWidth(0) = 700
.ColWidth(1) = 2300
.ColWidth(2) = 2300
.ColWidth(3) = 2000
'.ColHidden(0) = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Function StrList(cString)
Dim listTable As New adodb.Recordset
listTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until listTable.EOF
    StrList = StrList & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Private Sub xDesca_Change()
myLoad
End Sub
Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then myLoad
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
myLoad
End Sub
