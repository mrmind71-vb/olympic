VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StoreMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „Œ“‰"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
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
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   8460
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   6675
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   270
         TabIndex        =   8
         Top             =   585
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
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
         Height          =   195
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰"
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
         Height          =   195
         Left            =   5430
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   765
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·’‰›"
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
         Height          =   195
         Left            =   5295
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   3315
      End
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   945
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   375
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid invGrid 
      Height          =   7620
      Left            =   270
      TabIndex        =   9
      Top             =   1575
      Width           =   14865
      _cx             =   26220
      _cy             =   13441
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
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   9180
      Width           =   3210
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Œ—ÊÃ "
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton CmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈ŸÂ«— «·Õ—ﬂ…"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   1515
      End
   End
End
Attribute VB_Name = "StoreMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemTable As Recordset, StockTable As Recordset
Dim movetable, InvTable As Recordset
Dim searchTable, storeTable As Recordset
Dim SuppTable As Recordset
Dim CustTable As Recordset
Dim formMode, formFlag As Byte, nBal As Variant
Sub fillgrd()
nPrevious = 0
invGrid.Rows = 1
i = 1
cStockStr = "Item = " & MyParn(xItem.Text) & " And Store = " & MyParn(xStore.BoundText)

If InvTable.RecordCount > 0 Then
    InvTable.MoveFirst
    Do
       nPrevious = nPrevious + TurnValue(InvTable.In, Null, 0) - TurnValue(InvTable.out, Null, 0)
       invGrid.AddItem ""
       invGrid.TextArray(faIndex(i, 0, invGrid)) = InvTable.DESCA
       invGrid.TextArray(faIndex(i, 1, invGrid)) = TurnValue(InvTable.out, Null, 0)
       invGrid.TextArray(faIndex(i, 2, invGrid)) = TurnValue(InvTable.In, Null, 0)
       invGrid.TextArray(faIndex(i, 3, invGrid)) = TurnValue(nPrevious, Null, 0)
       invGrid.TextArray(faIndex(i, 5, invGrid)) = Format(InvTable.Date, "DD-MM-YYYY")
       invGrid.TextArray(faIndex(i, 4, invGrid)) = Format(InvTable.price, "#0.00")
       invGrid.TextArray(faIndex(i, 6, invGrid)) = TurnValue(InvTable.Doc_Id, Null, "")
       If InvTable!Type = "2" Or InvTable!Type = "7" Then
        SuppTable.FindFirst " CODE = " & MyParn(InvTable.CODE)
        If Not SuppTable.NoMatch Then invGrid.TextArray(faIndex(i, 7, invGrid)) = TurnValue(SuppTable.DESCA, Null, "")
       ElseIf InvTable!Type = "3" Or InvTable!Type = "6" Then
        CustTable.FindFirst " CODE = " & MyParn(InvTable.CODE)
        If Not CustTable.NoMatch Then invGrid.TextArray(faIndex(i, 7, invGrid)) = TurnValue(CustTable.DESCA, Null, "")
       End If
       InvTable.MoveNext
       i = i + 1
    Loop Until InvTable.EOF
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.grid1, 0)
Unload Search
End Sub
Function MYVALID()
MYVALID = False
If xItem.Text = "" Then Exit Function
ItemTable.FindFirst " item = " & MyParn(xItem.Text)
If ItemTable.NoMatch Then Exit Function
MYVALID = True
End Function

Private Sub cmdcorect_Click()

End Sub

Private Sub CmdGo_Click()
Dim cQryStr As String
ItemTable.FindFirst " item = " & MyParn(xItem.Text)
If ItemTable.NoMatch Then Exit Sub
'storeTable.FindFirst "Code = " & MyParn(xStore.BoundText)
'If storeTable.NoMatch Then Exit Sub
'CmdGo.Enabled = False

cQryStr = "Select * from File1_11 Where item = " & MyParn(xItem.Text)
If xStore.BoundText <> "" Then cQryStr = cQryStr & " and STORE = " & MyParn(xStore.BoundText)
cQryStr = cQryStr & " Order by [Date] "
Set InvTable = mydb.CreateSnapshot(cQryStr)
nBal = 0
If InvTable.RecordCount > 0 Then
InvTable.MoveFirst
Do
    nBal = nBal + TurnValue(InvTable.In, Null, 0) - TurnValue(InvTable.out, Null, 0)
    InvTable.MoveNext
Loop Until InvTable.EOF
End If
xBal.Caption = nBal
'x1.Caption = Val(Int(Val(XBAL.Caption) / ItemTable.Pack))
'x2.Caption = Val(XBAL.Caption) - (Val(x1.Caption) * ItemTable.Pack)
fillgrd
End Sub
Private Sub cmdExit_Click()
Unload StoreMove
End Sub
Private Sub CmdCreat_Click()
Me.MousePointer = 11
mydb.Execute "DELETE * FROM FILE1_11"

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,Out,DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'6'," & _
           "item,Quant," & _
           " '„»Ì⁄«  ' , " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_20 " & _
           " WHERE FILE6_20.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[in],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'3'," & _
           " item,Quant," & _
           " '„— Ã⁄«  ' , " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_10 " & _
           " WHERE FILE6_10.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[in],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'2'," & _
           "item,Quant," & _
           " '„‘ —Ì«  ' , " & _
           " Price,PRICE * Quant ,Store" & _
           " From File7_20 " & _
           " WHERE FILE7_20.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,Code,[Type],item,[out],DescA,Price,Total,store)" & _
           " Select [Date],Doc_No,Code,'7'," & _
           "item,Quant," & _
           " '„—œÊœ „‘ —Ì«   ' , " & _
           " Price,PRICE * Quant ,Store" & _
           " From File6_11 " & _
           " WHERE FILE6_11.STORE <> 'SS'  "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [Date],Doc_No,'8'," & _
           "item,QUANT," & _
           " '’«œ— ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           "STORE" & _
           " From file1_81 "
'mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[IN],DescA,STORE)" & _
           " Select [Date],Doc_No,'4'," & _
           "item,QUANT," & _
           " 'Ê«—œ ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           " STORE" & _
           " From file1_80 "
'mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [Date],Doc_No,'9'," & _
           "item,QUANT," & _
           " 'Â«·ﬂ ›Ï' & Format([Date], 'dd-mm-yy'), " & _
           "STORE" & _
           " From file1_82 "
'mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[Out],DescA,STORE)" & _
           " Select [DATE],Doc_No,'F'," & _
           "item,QUANT," & _
           " ' ÕÊÌ·«  ≈·Ï' & STORE2, " & _
           "STORE1" & _
           " From file1_60 "
mydb.Execute cString

cString = "INSERT INTO FILE1_11( " & _
           "[Date],Doc_Id,[Type],item,[IN],DescA,STORE)" & _
           " Select [Date],Doc_No, 'T'," & _
           "item,QUANT," & _
           " ' ÕÊÌ·«  „‰' & STORE1, " & _
           "STORE2" & _
           " From file1_60 "
mydb.Execute cString


cString = "insert into File1_11(type,item,[date],store,desca,Doc_Id,[In])" & _
        " Select 'z',item,[date],store,' Ã—œ ›Ï ' & Format(Date,'dd-mm-yyyy'),Doc_NO,differ From File0_10 where Closed"
'mydb.Execute cString
Me.MousePointer = 1


End Sub

Private Sub Form_Load()
Set ItemTable = mydb.CreateSnapshot("file1_10")
Set storeTable = mydb.CreateSnapshot("Stores")

Data1.DatabaseName = MdbPath
Data1.RecordSource = "Stores"
xStore.BoundColumn = "Code"
xStore.ListField = "DescA"
With invGrid
invGrid.Cols = 8
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = "’«œ—"
.TextMatrix(0, 2) = "Ê«—œ"
.TextMatrix(0, 3) = "—’Ìœ"
.TextMatrix(0, 4) = "”⁄—"
.TextMatrix(0, 5) = " «—ÌŒ"
.TextMatrix(0, 6) = "„” ‰œ"
.TextMatrix(0, 7) = "≈”„"


invGrid.ColWidth(0) = 2000
invGrid.ColWidth(1) = 1200
invGrid.ColWidth(2) = 1200
invGrid.ColWidth(3) = 1200
invGrid.ColWidth(4) = 1200
invGrid.ColWidth(5) = 1200
invGrid.ColWidth(6) = 1200
invGrid.ColWidth(7) = 2500
End With
For i = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(i) = 1
Next
End Sub

Private Sub xBal_Click()

End Sub

Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then ItemsLookup
End Sub
Private Sub XITEM_LostFocus()
If Not CmdGo.Enabled And xStore.BoundText <> "" Then CmdGo.Enabled = True
ItemTable.FindFirst "item = " & MyParn(xItem.Text)
If ItemTable.NoMatch Then
    xDesca.Caption = ""
    Exit Sub
End If
xDesca.Caption = ItemTable.DESCA
End Sub
Private Sub xStore_Click(Area As Integer)
If Not CmdGo.Enabled Then CmdGo.Enabled = True
End Sub
