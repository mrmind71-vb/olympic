VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chargecodefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«ﬂÊ«œ «·„’«—Ì› "
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6120
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   90
      TabIndex        =   2
      Top             =   3735
      Width           =   5910
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   315
         TabIndex        =   4
         Top             =   540
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "«·Ê’› :"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   780
      End
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
      Height          =   420
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   3690
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5910
      _cx             =   10425
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
      Cols            =   3
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
      TabBehavior     =   1
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
Attribute VB_Name = "chargecodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Public myPublic As Integer, bedit As Boolean
Dim GRDTABLE As New adodb.Recordset
Dim CTABLE As String, CGROUP As String
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command2_Click()
myLoad
End Sub
Private Sub myLoad()
Dim cFilter As String
GRDTABLE.Requery
If Trim(xname.Text) <> "" Then cFilter = MyParnAnd(xname.Text, "desca")
If Trim(xGroup.BoundText) <> "" Then cFilter = cFilter & turnFound2(cFilter, "") & " [GROUP] = " & MyParn(xGroup.BoundText)
GRDTABLE.Filter = cFilter
i = 0
grid1.Rows = 1
Do Until GRDTABLE.EOF
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = GRDTABLE!CODE & ""
    grid1.TextMatrix(grid1.Rows - 1, 1) = GRDTABLE!desca & ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = GRDTABLE!Group & ""
    GRDTABLE.MoveNext
Loop
grid1.AddItem ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GRDTABLE.Close
Set GRDTABLE = Nothing
Set chargecodefrm = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
con.BeginTrans
If grid1.TextMatrix(Row, 0) <> "" Then
    con.Execute "update " & CTABLE & _
                " set DESCA = " & addstring(grid1.TextMatrix(Row, 1)) & "," & _
                " [GROUP] = " & addstring(grid1.TextMatrix(Row, 2)) & _
                " WHERE CODE = " & MyParn(grid1.TextMatrix(Row, 0))
Else
    If Trim(grid1.TextMatrix(Row, 1)) <> "" Then
        grid1.TextMatrix(Row, 0) = RetZero(Newflag(CTABLE, "code"), 3)
        For nTry = 1 To 10
            con.Execute "INSERT INTO " & CTABLE & "(CODE,DESCA,[GROUP]) " & _
                        "VALUES( " & _
                        addstring(grid1.TextMatrix(Row, 0)) & "," & _
                        addstring(grid1.TextMatrix(Row, 1)) & "," & _
                        addstring(grid1.TextMatrix(Row, 2)) & _
                        ")"
            
            If Err.Number = -2147467259 And nTry < 10 Then
                Err.Clear
                grid1.TextMatrix(Row, 0) = RetZero(Val(grid1.TextMatrix(Row, 0)) + 1, 3)
            End If
            If Err.Number = 0 Then Exit For
            If Err.Number <> 0 Then GoTo myerror
        Next
    End If
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
If KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Remove Row ?? Are you sure", vbYesNo) = vbYes Then
            con.BeginTrans
            con.Execute "Delete  from " & CTABLE & " where code = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.Row
            grid1_EnterCell
        End If
    End If
End If
End Sub
Private Sub Form_Load()
openCon con
If myPublic = 1 Then
    CTABLE = "file8_51"
    CGROUP = "file8_52"
    Me.Caption = "«ﬂÊ«œ «·„’«—Ì›"
Else
    CTABLE = "file8_61"
    CGROUP = "file8_62"
    Me.Caption = "«ﬂÊ«œ «·«Ì—«œ« "
End If

data1.ConnectionString = strCon
data1.RecordSource = CGROUP
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

cString = "SELECT code,desca,[GROUP] from " & CTABLE & " order by code"
GRDTABLE.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With grid1

.ColWidth(0) = 1000
.ColWidth(1) = 2000
.ColWidth(2) = 2000
.TextMatrix(0, 0) = "«·ﬂÊœ"
.TextMatrix(0, 1) = "«·»Ì«‰"
.TextMatrix(0, 2) = "«·„Ã„Ê⁄…"
.ColAlignment(1) = flexAlignRightCenter
.ColAlignment(2) = flexAlignRightCenter
.ColComboList(2) = StrList
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
Private Sub xcountry_code_LostFocus()
xcountry.Text = ""
If Trim(xcountry.Text) <> "" Then xcountry.Text = GetDesca("select country_name from country where country_code = " & Val(xcountry_code.Text))
myLoad
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "description required"
        Cancel = True
    End If
End If
End Sub

Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then myLoad
End Sub

Private Sub xname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myLoad
End Sub
Private Function StrList()
Dim listTable As New adodb.Recordset
listTable.Open "SELECT * FROM " & CGROUP & " ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
StrList = "#  " & ";       "
Do Until listTable.EOF
    StrList = StrList & "|#" & listTable!CODE & ";" & listTable!desca
    listTable.MoveNext
Loop
End Function

