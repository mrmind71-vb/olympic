VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form itemsgrdfrm 
   Caption         =   "»Ì«‰«  «·«’‰«›"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2295
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4455
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   1125
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
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
      Height          =   8520
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   14910
      _cx             =   26300
      _cy             =   15028
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
      BackColorSel    =   8454143
      ForeColorSel    =   128
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   6
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
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   2295
      TabIndex        =   5
      Top             =   8595
      Width           =   12750
      Begin VB.CommandButton cmdSection 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   495
         Width           =   465
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   135
         Width           =   465
      End
      Begin VB.TextBox XITEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   4065
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   4065
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   9225
         TabIndex        =   2
         Top             =   495
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   9225
         TabIndex        =   1
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11745
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   495
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   765
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8595
      Width           =   1500
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "itemsgrd2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
   End
End
Attribute VB_Name = "itemsgrdfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bedit As Boolean
Dim clist1 As String, cList2 As String, cList3 As String, cList4 As String
Dim CardTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub myload()
On Error GoTo myerror
'                   0                   1                           2                       3                          4                    5                    6                       7
cString = "SELECT ITEM as [«·ﬂÊœ],FILE1_10.[GROUP] as [«·„Ã„Ê⁄…],[SECTION] as [«·ﬁ”„],FILE1_10.DESCA as [«·»Ì«‰],FILE1_10.COST AS [«· ﬂ·›…],FILE1_10.MONTHES AS [«·„œ… »«·‘Â—],KILO AS [ﬂÌ·Ê „ —],FILE1_10.ITEM  " & _
          " FROM FILE1_10"

If IsNumeric(xGroup.BoundText) Then
    cString = cString & turn(cString) & "FILE1_10.[GROUP] = " & xGroup.BoundText
End If

If IsNumeric(xSection.BoundText) Then
    cString = cString & turn(cString) & "FILE1_10.[SECTION] = " & xSection.BoundText
End If

If Trim(xDesca.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDesca.Text, "FILE1_10.desca")
End If

If Trim(XITEM.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(XITEM.Text, "ITEM")
End If

cString = cString & " order by FILE1_10.ID_ITEM "
DATA1.RecordSource = cString
DATA1.Refresh
myAddItem
'lblTotal.Caption = IIf(grid1.Rows < 3, "", "≈Ã„«·Ì ⁄œœ «·«’‰«› : " & grid1.Rows - 2)
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExel_Click()
grid1.ColHidden(2) = True
grid1.ColHidden(3) = True
grid1.ColHidden(4) = True
ToFileExel grid1
grid1.ColHidden(2) = False
grid1.ColHidden(3) = False
grid1.ColHidden(4) = False
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGroup_Click()
Dim oFlagGroup As New FlagGroupFrm
oFlagGroup.sCaption = "„Ã„Ê⁄«  «·«’‰«›"
oFlagGroup.sCode = "«·ﬂÊœ"
oFlagGroup.sDesca = "≈”„ «·„Ã„Ê⁄…"
oFlagGroup.sGroupDesca = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
oFlagGroup.sTable = "FILE1_50"
oFlagGroup.sTableGroup = "FILE1_50G"
oFlagGroup.nZero = -1
oFlagGroup.nZeroGroup = -1
oFlagGroup.sGroupCaption = "„Ã„Ê⁄«  «·«’‰«› «·—∆Ì”Ì…"
oFlagGroup.Show 1
clist1 = StrList2("select * from file1_50 order by desca")
grid1.ColComboList(2) = clist1
DATA2.Refresh
End Sub

Private Sub cmdSection_Click()
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sTable = "FILE1_10SC"
oFlagfrm.sCaption = "«ﬁ”«„ «·«’‰«›"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
cList2 = StrList2("select * from file1_10SC order by desca")
grid1.ColComboList(2) = cList2
DATA3.Refresh
End Sub

Private Sub Command2_Click()
'If grid1.Rows = 1 Then
'    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
'    Exit Sub
'End If
'
'If Not doprint Then
'    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
'    Exit Sub
'End If
'CardPrintNew1.PrintArray
'CardPrintNew1.Show 1
'getItems
'getItems2
transDatafrm.Show 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set itemsgrdfrm = Nothing
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
If grid1.Row = grid1.Rows - 1 Then myAddItem
If myreplace(Row) Then
   If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then myload
End If
End Sub

Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
    If (grid1.Col = 0) Then
        grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDKbdMouse
    End If
End With
End Sub

Private Sub Grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub Form_Load()
openCon con

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE1_50"
Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "FILE1_10SC"
Set xSection.RowSource = DATA3
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"


Set grid1.DataSource = DATA1
DATA1.ConnectionString = strCon
With grid1
clist1 = StrList2("Select code,desca from file1_50 order by desca")
cList2 = StrList2("Select code,desca from file1_10sc order by desca")

myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem OldRow
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "‰Ê⁄ «·’‰› „ÿ·Ê»"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬁ”„ «·’‰› „ÿ·Ê»"
        Cancel = True
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.ColComboList(1) = clist1
.ColComboList(2) = cList2
.ColWidth(0) = 1000
.ColWidth(1) = 2500
.ColWidth(2) = 2000
.ColWidth(3) = 3000
.ColWidth(4) = 2000
.ColWidth(5) = 1500
.ColWidth(6) = 1000
.ColHidden(.Cols - 1) = True
.RowHeight(0) = 800
.WordWrap = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
If grid1.Rows > 1 Then .Cell(flexcpFontSize, 1, 1, .Rows - 1, 1) = 10
End With
End Sub

Private Sub xDesca_Change()
    myload
End Sub
Private Sub xGroup_Change()
If xGroup.MatchedWithList Or Trim(xGroup.BoundText) = "" Then
    myload
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 0
End If
End Sub
Private Sub xITEM_Change()
myload
End Sub
Private Sub xSection_Change()
If xSection.MatchedWithList Or Trim(xSection.BoundText) = "" Then
    myload
    grid1.Select grid1.Rows - 1, 1
    grid1.ShowCell grid1.Rows - 1, 0
End If
End Sub
Private Sub xSection_LostFocus()
With xSection
If Not .MatchedWithList Then .BoundText = ""
End With
End Sub
Private Function validRow(Row) As Boolean
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(grid1.TextMatrix(Row, 2)) = "" Then Exit Function
If Trim(grid1.TextMatrix(Row, 3)) = "" Then Exit Function
validRow = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–› «·’‰› !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
            con.BeginTrans
            con.Execute "delete from file1_10 where item = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.Row
        End If
    End If
ElseIf KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
con.RollbackTrans
myload
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Not (Col = 1 Or Col = 2) Then CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    If Col = 0 And grid1.TextMatrix(Row, 1) <> "" Then
          grid1.Select Row, 2 + IIf(grid1.TextMatrix(grid1.Rows - 1, 2) <> "", 1, 0)
    ElseIf Col = 1 And grid1.TextMatrix(Row, 2) <> "" Then
        grid1.Select Row, 3
    Else
        grid1.Select Row, Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = Row + 1
    grid1.Select Row + 1, 1 + IIf(grid1.TextMatrix(grid1.Rows - 1, 1) <> "", 1, 0) + IIf(grid1.TextMatrix(grid1.Rows - 1, 1) <> "" And grid1.TextMatrix(grid1.Rows - 1, 2) <> "", 1, 0)
    grid1.ShowCell Row + 1, 1 + IIf(grid1.TextMatrix(grid1.Rows - 1, 1) <> "", 1, 0) + IIf(grid1.TextMatrix(grid1.Rows - 1, 1) <> "" And grid1.TextMatrix(grid1.Rows - 1, 2) <> "", 1, 0)
End If
End Sub
Private Sub myAddItem()
With grid1
    .AddItem ""
    If .Rows > 2 Then
        grid1.TextMatrix(.Rows - 1, 2 - 1) = grid1.TextMatrix(.Rows - 2, 2 - 1)
        grid1.TextMatrix(.Rows - 1, 3 - 1) = grid1.TextMatrix(.Rows - 2, 3 - 1)
    End If
    If xGroup.MatchedWithList And grid1.TextMatrix(.Rows - 1, 2 - 1) = "" Then
        grid1.TextMatrix(.Rows - 1, 2 - 1) = xGroup.BoundText
    End If
    If xSection.MatchedWithList And grid1.TextMatrix(.Rows - 1, 3 - 1) = "" Then
        grid1.TextMatrix(.Rows - 1, 3 - 1) = xSection.BoundText
    End If
End With
End Sub
Private Function myreplace(Row As Long) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[GROUP]", addvalue(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[SECTION]", addvalue(grid1.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(grid1.TextMatrix(Row, 3)))
aInsert = AddFlag(aInsert, "[COST]", Val(grid1.TextMatrix(Row, 4)))
aInsert = AddFlag(aInsert, "[MONTHES]", Val(grid1.TextMatrix(Row, 5)))
aInsert = AddFlag(aInsert, "[KILO]", Val(grid1.TextMatrix(Row, 6)))
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    Dim sItem As String
    sItem = Newflag("file1_10", "CONVERT(INT,ITEM)", con)
    If Val(sItem) < 101 Then sItem = 101
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(sItem))
    con.Execute addInsert(aInsert, "FILE1_10")
Else
    con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.item = " & MyParn(grid1.TextMatrix(Row, grid1.Cols - 1)))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Function
Private Sub xGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    grid1.SetFocus
    CellPos KeyCode, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub
Private Sub xSection_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    grid1.SetFocus
    CellPos KeyCode, grid1.Rows - 2, grid1.Cols - 1
End If
End Sub

