VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmProft 
   Caption         =   "Õ”«» «—»«Õ Ê Œ”«∆—"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_print 
      Caption         =   "ÿ»«⁄…"
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
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   900
      Width           =   1455
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "⁄—÷"
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
      Left            =   1035
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   450
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6270
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   2490
      End
      Begin VB.TextBox xdate1 
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
         Height          =   330
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   2490
      End
      Begin MSDataListLib.DataCombo xStore 
         Bindings        =   "FrmProft.frx":0000
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   900
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Õ· :"
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
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
   End
   Begin TabDlg.SSTab stab 
      Height          =   4905
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   8652
      _Version        =   393216
      MousePointer    =   8
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Õ”«» «—»«Õ ÊŒ”«∆—"
      TabPicture(0)   =   "FrmProft.frx":0014
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "VSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Õ”«» «·„ «Ã—…"
      TabPicture(1)   =   "FrmProft.frx":0030
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4365
         Left            =   90
         TabIndex        =   1
         Top             =   405
         Width           =   8655
         _cx             =   15266
         _cy             =   7699
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
      Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   4365
         Left            =   -74910
         TabIndex        =   9
         Top             =   405
         Width           =   8655
         _cx             =   15266
         _cy             =   7699
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
         FixedRows       =   0
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
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   180
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmProft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PurchTable As Recordset
Dim SalesTable As Recordset
Dim RPurchTable As Recordset
Dim TransInTable As Recordset
Dim TransOutTable As Recordset
Dim FBalTable As Recordset
Dim CBalTable As Recordset
Dim ChargTable As Recordset
Dim Charg2Table As Recordset
Dim VisaTable As Recordset
Dim Disc1Table As Recordset
Dim DISC2Table As Recordset
Private Sub CMD_PRINT_Click()
Dim temptable As Recordset
tempdb.Execute " DELETE * FROM TEMP "

Set temptable = tempdb.OpenRecordset(" SELECT * FROM TEMP ")
With VsProft
    For I = 1 To .Rows - 1
        temptable.AddNew
        temptable.str1 = TurnValue(.TextMatrix(I, 0), "", Null)
        temptable.str2 = TurnValue(.TextMatrix(I, 2), "", Null)
        temptable.val1 = Val(.TextMatrix(I, 1))
        temptable.val2 = Val(.TextMatrix(I, 3))
        temptable.str18 = "Õ”«» „ «Ã—… „‰  «—ÌŒ " & xDate1.Text & "  ≈·Ï  «—ÌŒ " & xDate2.Text
        temptable.STR19 = TurnValue(Me.xstore.Text, "", Null)
        temptable.Update
    Next I
End With

myws.BeginTrans
myws.CommitTrans

Report1.ReportFileName = PublicPath & "\Reports\RProft_1.rpt"
Report1.DataFiles(0) = tempPath
Report1.Action = 1


tempdb.Execute " DELETE * FROM TEMP "
Set temptable = tempdb.OpenRecordset(" SELECT * FROM TEMP ")
With VsProft2
    For I = 1 To .Rows - 1
        temptable.AddNew
        temptable.str1 = TurnValue(.TextMatrix(I, 0), "", Null)
        temptable.str2 = TurnValue(.TextMatrix(I, 2), "", Null)
        temptable.val1 = Val(.TextMatrix(I, 1))
        temptable.val2 = Val(.TextMatrix(I, 3))
        temptable.str18 = "Õ”«» √—»«Õ Ê Œ”«∆— „‰  «—ÌŒ " & xDate1.Text & "  ≈·Ï  «—ÌŒ " & xDate2.Text
        temptable.STR19 = TurnValue(Me.xstore.Text, "", Null)
        temptable.Update
    Next I
End With
myws.BeginTrans
myws.CommitTrans

Report1.ReportFileName = PublicPath & "\Reports\RProft_2.rpt"
Report1.DataFiles(0) = tempPath
Report1.Action = 1
End Sub
Private Sub Form_Load()
data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & MdbPath
data1.RecordSource = "stores"
data1.Refresh
xstore.ListField = "DESCA"
xstore.BoundColumn = "CODE"
data1.Refresh

With VsProft
    .Rows = 7
    For I = 0 To .Rows - 1
        .RowHeight(I) = 300
    Next I
    .Cols = 4
    .FixedCols = 0
    .FixedRows = 0
    .ColWidth(0) = 1500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
    .TextMatrix(1, 0) = "—’Ìœ √Ê·"
    .TextMatrix(1, 2) = "—’Ìœ √Œ—"
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .TextMatrix(0, 0) = "„œÌ‰"
    .TextMatrix(0, 1) = "„œÌ‰"
    .TextMatrix(0, 2) = "œ«∆‰"
    .TextMatrix(0, 3) = "œ«∆‰"
    
    .TextMatrix(2, 0) = "„‘ —Ì« "
    .TextMatrix(2, 2) = "„»Ì⁄« "
'    .TextMatrix(3, 2) = "Œ’„ ›Ì“«"
'    .TextMatrix(4, 0) = "Œ’„ „»Ì⁄« "
    .TextMatrix(3, 2) = " ”ÊÌ… „Ê—œÌ‰"
    .RowHidden(4) = True
    .TextMatrix(5, 0) = " ÕÊÌ·«  ≈·‹Ï"
    .TextMatrix(5, 2) = " ÕÊÌ·«  „‹‰"
    
    .Row = 0: .Col = 0
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = 0: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = 0: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    .Row = 0: .Col = 3
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    
    .Row = 1: .Col = 0
    .CellBorder vbbrown, 0, 1, 0, 0, 1, 1
    .Row = 1: .Col = 1
    .CellBorder vbbrown, 1, 1, 0, 0, 1, 1
    .Row = 1: .Col = 2
    .CellBorder vbbrown, 0, 1, 1, 0, 1, 1
    .Row = 1: .Col = 3
    .CellBorder vbbrown, 0, 1, 0, 0, 1, 1
    
    .Row = 2: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 0, 1, 1
    .Row = 3: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 0, 1, 1
    .Row = 4: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 0, 1, 1
    .Row = 5: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 0, 1, 1
    .Row = 6: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 0, 1, 1

    .Row = 2: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
    .Row = 3: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
    .Row = 4: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
    .Row = 5: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
    .Row = 6: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7

End With


With VsProft2
    .Rows = 2
    For I = 0 To .Rows - 1
        .RowHeight(I) = 300
    Next I
    .Cols = 4
    .FixedCols = 0
    .FixedRows = 0
    .ColWidth(0) = 1500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .TextMatrix(0, 0) = "„œÌ‰"
    .TextMatrix(0, 1) = "„œÌ‰"
    .TextMatrix(0, 2) = "œ«∆‰"
    .TextMatrix(0, 3) = "œ«∆‰"
    .Row = 0: .Col = 0
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = 0: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    
    .Row = 0: .Col = 0
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = 0: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = 0: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    .Row = 0: .Col = 3
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    
    .Row = 1: .Col = 0
    .CellBorder vbbrown, 0, 1, 0, 0, 1, 1
    .Row = 1: .Col = 1
    .CellBorder vbbrown, 1, 1, 0, 0, 1, 1
    .Row = 1: .Col = 2
    .CellBorder vbbrown, 0, 1, 1, 0, 1, 1
    .Row = 1: .Col = 3
    .CellBorder vbbrown, 0, 1, 0, 0, 1, 1

End With

End Sub
Private Sub Cmd_OK_Click()
    Dim nT1, nT2, nCharg As Double
    Dim nProf, nDameg As Double
    Dim nTT1, nTT2 As Double
    nT1 = 0
    nT2 = 0
    If Not IsDate(xDate1.Text) Or Not IsDate(xDate2.Text) Then Exit Sub
    FixCost
    
    cStr1 = " SELECT Sum([IN]*[FILE1_10].[COST]) AS TIN, Sum([OUT]*FILE1_10.[COST]) AS TOUT FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
            " WHERE DATE < " & DateSql(xDate1.Text)
    
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(xstore.BoundText)
    Set FBalTable = mydb.OpenRecordset(cStr1)
    
    mydb.Execute " UPDATE FILE1_10 LEFT JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM SET FILE1_10.cost = [FILE7_20].[PRICE] WHERE DATE <= " & DateSql(xDate2.Text)
    
    cStr1 = " SELECT Sum(FILE7_20.TOTAL) AS T_TOTAL FROM FILE7_20 " & _
            " WHERE DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(xstore.BoundText)
    Set PurchTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum(FILE6_11.TOTAL) AS T_TOTAL FROM FILE6_11 " & _
            " WHERE DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(xstore.BoundText)
    Set RPurchTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum(FILE6_20.TOTAL) AS T_TOTAL FROM FILE6_20 " & _
            " WHERE DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(xstore.BoundText)
    Set SaleTable = mydb.OpenRecordset(cStr1)


    cStr1 = " SELECT Sum([IN]*FILE1_10.[COST]) AS TIN, Sum([OUT]*FILE1_10.[COST]) AS TOUT FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
            " WHERE DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(xstore.BoundText)
    Set CBalTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum([QUANT]*FILE1_10.[COST]) AS TCost FROM FILE1_10 INNER JOIN FILE1_60 ON FILE1_10.ITEM = FILE1_60.ITEM " & _
            " WHERE DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE2 = " & MyParn(xstore.BoundText)
    Set TransInTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT Sum([QUANT]*FILE1_10.[COST]) AS TCost FROM FILE1_10 INNER JOIN FILE1_60 ON FILE1_10.ITEM = FILE1_60.ITEM " & _
            " WHERE DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND STORE1 = " & MyParn(xstore.BoundText)
    Set TransOutTable = mydb.OpenRecordset(cStr1)

    cStr1 = " SELECT FILE1_70.DESCA, Sum(FILE8_50.VALUE) AS tVALUE " & _
            " FROM (FILE1_70 RIGHT JOIN FILE8_70 ON FILE1_70.CODE = FILE8_70.MAINGROUP) RIGHT JOIN FILE8_50 ON FILE8_70.CODE = FILE8_50.CHARGE WHERE FILE1_70.FLAG = 7 " & _
            " AND DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    If xstore.BoundText <> "" Then cStr1 = cStr1 & " AND BOX = " & MyParn(xstore.BoundText)
    cStr1 = cStr1 & " GROUP BY FILE1_70.DESCA "
    Set ChargTable = mydb.OpenRecordset(cStr1)


    cStr1 = " SELECT Sum(FILE8_20.value) AS TDISC FROM FILE8_20 WHERE CODE IS NOT NULL " & _
            " AND DATE >= " & DateSql(xDate1.Text) & _
            " AND DATE <= " & DateSql(xDate2.Text)
    Set DISC2Table = mydb.OpenRecordset(cStr1)
    
    With VsProft
    .Rows = 7
    For I = 1 To .Rows - 1
        .TextMatrix(I, 1) = ""
        .TextMatrix(I, 3) = ""
    Next I
    If PurchTable.RecordCount > 0 Then
        PurchTable.MoveFirst
        .TextMatrix(2, 1) = TurnValue(PurchTable.T_TOTAL, Null, 0)
    End If
    
    If RPurchTable.RecordCount > 0 Then
        RPurchTable.MoveFirst
        .TextMatrix(2, 1) = Val(.TextMatrix(2, 1)) - TurnValue(RPurchTable.T_TOTAL, Null, 0)
    End If
    
    If SaleTable.RecordCount > 0 Then
        SaleTable.MoveFirst
        .TextMatrix(2, 3) = TurnValue(SaleTable.T_TOTAL, Null, 0)
    End If
    
    If FBalTable.RecordCount > 0 Then
        FBalTable.MoveFirst
        .TextMatrix(1, 1) = TurnValue(FBalTable.TIN, Null, 0) - TurnValue(FBalTable.TOUT, Null, 0)
    End If
    
    If CBalTable.RecordCount > 0 Then
        CBalTable.MoveFirst
        .TextMatrix(1, 3) = TurnValue(CBalTable.TIN, Null, 0) - TurnValue(CBalTable.TOUT, Null, 0)
    End If
    
    
    If xstore.BoundText = "" Then
'        If Disc1Table.RecordCount > 0 Then
'            Disc1Table.MoveFirst
'            .TextMatrix(4, 1) = TurnValue(Disc1Table.tdisc, Null, 0)
'        End If
    
        If DISC2Table.RecordCount > 0 Then
            DISC2Table.MoveFirst
            .TextMatrix(3, 3) = TurnValue(DISC2Table.Tdisc, Null, 0)
        End If
    End If
    
    If xstore.BoundText <> "" Then
        If TransInTable.RecordCount > 0 Then
            TransInTable.MoveFirst
            .TextMatrix(5, 1) = TurnValue(TransInTable.Tcost, Null, 0)
            .TextMatrix(5, 0) = " ÕÊÌ·«  ≈·Ï"
        End If
        If TransOutTable.RecordCount > 0 Then
            TransOutTable.MoveFirst
            .TextMatrix(5, 3) = TurnValue(TransOutTable.Tcost, Null, 0)
            .TextMatrix(5, 2) = " ÕÊÌ·«  ≈·Ï"
        End If
    End If
    
    For I = 1 To 5
        nT1 = nT1 + Val(.TextMatrix(I, 1))
        nT2 = nT2 + Val(.TextMatrix(I, 3))
        .TextMatrix(I, 1) = Format(.TextMatrix(I, 1), "#0.00")
        .TextMatrix(I, 3) = Format(.TextMatrix(I, 3), "#0.00")
    Next I
    If nT1 > nT2 Then
        .TextMatrix(6, 3) = Format(nT1 - nT2, "#0.00")
        .TextMatrix(6, 2) = "Œ”«—…"
        nDameg = nT1 - nT2
    End If
    If nT1 < nT2 Then
        .TextMatrix(6, 1) = Format(nT2 - nT1, "#0.00")
        .TextMatrix(6, 0) = "—»‹‹‹Õ"
        nProft = nT2 - nT1
    End If
    
    nTT1 = 0
    nTT2 = 0
    For I = 1 To .Rows - 1
        nTT1 = nTT1 + Val(.TextMatrix(I, 1))
        nTT2 = nTT2 + Val(.TextMatrix(I, 3))
    Next I
    .AddItem ""
    .TextMatrix(.Rows - 1, 1) = Format(nTT1, "#0.00")
    .TextMatrix(.Rows - 1, 3) = Format(nTT2, "#0.00")
    
    .Row = .Rows - 2: .Col = 0
    .CellBorder vbbrown, 0, 0, 0, 1, 1, 1
    .Row = .Rows - 2: .Col = 1
    .CellBorder vbbrown, 1, 0, 0, 1, 1, 1
    .Row = .Rows - 2: .Col = 2
    .CellBorder vbbrown, 0, 0, 1, 1, 1, 1
    .Row = .Rows - 2: .Col = 3
    .CellBorder vbbrown, 0, 0, 0, 1, 1, 1
    
    .Row = .Rows - 1: .Col = 0
    .CellBorder vbbrown, 0, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 1
    .CellBorder vbbrown, 1, 1, 0, 1, 1, 1
    .Row = .Rows - 1: .Col = 2
    .CellBorder vbbrown, 0, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 3
    .CellBorder vbbrown, 1, 1, 0, 1, 1, 1
    
    End With
        
    With VsProft2
    .Rows = 2
    .TextMatrix(1, 0) = ""
    .TextMatrix(1, 1) = ""
    .TextMatrix(1, 2) = ""
    .TextMatrix(1, 3) = ""
  
    If nProft >= 0 Then
        .TextMatrix(1, 2) = "«·√—»«Õ"
        .TextMatrix(1, 3) = Format(nProft, "#0.00")
    Else
        .TextMatrix(1, 0) = "Œ”«—…"
        .TextMatrix(1, 1) = Format(nDameg, "#0.00")
    End If
    
    If ChargTable.RecordCount > 0 Then
        ChargTable.MoveFirst
        Do While True
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = ChargTable.desca
            .TextMatrix(.Rows - 1, 1) = Format(ChargTable.TValue, "#0.00")
            nCharg = nCharg + Val(.TextMatrix(.Rows - 1, 1))

            .Row = .Rows - 1: .Col = 1
           .CellBorder vbbrown, 1, 0, 0, 0, 1, 1
            .Row = .Rows - 1: .Col = 2
            .CellBorder vbbrown, 0, 0, 1, 0, 1, 1
            ChargTable.MoveNext
            If ChargTable.EOF Then Exit Do
       Loop
    End If
    
    
    .AddItem ""
    .Row = .Rows - 1: .Col = 0
    .CellBorder vbbrown, 1, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 1
    .CellBorder vbbrown, 1, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 2
    .CellBorder vbbrown, 1, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 3
    .CellBorder vbbrown, 1, 1, 1, 1, 1, 1
    If nProft >= 0 Then
        If nProft - nCharg > 0 Then
            .TextMatrix(.Rows - 1, 0) = "’«›Ï «·√—»«Õ"
            .TextMatrix(.Rows - 1, 1) = Format(nProft - nCharg, "#0.00")
        Else
            .TextMatrix(.Rows - 1, 2) = "’«›Ï «·Œ”«—…"
            .TextMatrix(.Rows - 1, 3) = Format((nProft - nCharg) * -1, "#0.00")
        End If
    Else
        .TextMatrix(.Rows - 1, 2) = "’«›Ï «·Œ”«—…"
        .TextMatrix(.Rows - 1, 3) = Format(nDameg + nCharg, "#0.00")
    End If
    
    nTT1 = 0
    nTT2 = 0
    For I = 1 To .Rows - 1
        nTT1 = nTT1 + Val(.TextMatrix(I, 1))
        nTT2 = nTT2 + Val(.TextMatrix(I, 3))
    Next I
    .AddItem ""
    .Row = .Rows - 1: .Col = 0
    .CellBorder vbbrown, 0, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 1
    .CellBorder vbbrown, 1, 1, 0, 1, 1, 1
    .Row = .Rows - 1: .Col = 2
    .CellBorder vbbrown, 0, 1, 1, 1, 1, 1
    .Row = .Rows - 1: .Col = 3
    .CellBorder vbbrown, 1, 1, 0, 1, 1, 1
    
    .TextMatrix(.Rows - 1, 1) = Format(nTT1, "#0.00")
    .TextMatrix(.Rows - 1, 3) = Format(nTT2, "#0.00")
        
    End With
End Sub
Private Sub FixCost()
Dim loctable As New ADODB.Recordset
loctable.Open "FILE1_10", CON, adOpenKeyset, adLockReadOnly, adCmdTable
On Error GoTo MyError
CON.BeginTrans
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
I = 0
prog1.Value = 0
prog1.Visible = True
Do Until loctable.EOF
    I = I + 1
    prog1.Value = Round(I / nRecordCount, 2) * 100
    CON.Execute "UPDATE FILE1_10 SET FILE1_10.COST = " & itemCost(loctable!Item) & " where item = " & MyParn(loctable!Item)
    loctable.MoveNext
Loop
CON.Execute "update (file1_10 inner join file7_20 on file1_10.item = file7_20.item) inner join file7_20h on file7_20.doc_no = file7_20h.doc_no set file1_10.supler = file7_20h.code"
CON.CommitTrans
prog1.Visible = False
MsgBox "DONE..."
Exit Sub
MyError:
MsgBox Err.Description
Err.Clear
CON.RollbackTrans
End Sub
Private Sub MyLoad1()
'cFirstBal = myiif("Date < " & DateSq(xdate1.Text), "(VAL([IN] & '') - VAL([OUT] & '')) * VAL(FILE1_10.COST & '')") & " as FirstBal"
'cLastBal = myiif("Date <= " & DateSq(xDate2.Text), "(VAL([IN] & '') - VAL([OUT] & '')) * VAL(FILE1_10.COST & '')") & " as LastBal"
'cFirstBal = myiif("Date < " & DateSq(xdate1.Text) & turnFound(Trim(xStore.BoundText), " and  store = " & MyParn(xStore.BoundText)), "(VAL([IN] & '') - VAL([OUT] & ''))") & " as FirstBal"
'cLastBal = myiif("Date <= " & DateSq(xDate2.Text) & turnFound(Trim(xStore.BoundText), " and  store = " & MyParn(xStore.BoundText)), "(VAL([IN] & '') - VAL([OUT] & '')) ") & " as LastBal"

If IsDate(xDate2.Text) Then
    cString = "Select Sum(Val([IN] & '') - VAL(OUT & '')) * VAL(FILE1_11.COST & '') FROM FILE1_11 WHERE DATE < " & DateSq(xDate1.Text)
    If xstore.BoundText <> "" Then
        cString = cString & turnFound(cString) & " STORE = " & MyParn(xstore.BoundText)
    End If
    nFirstBal = Val(GetDesca(cString))
End If


'„‘ —Ì« 

cField1 = "Select Sum((Val(File7_20.Quant & '') * Val(File7_20.Price & '')) + Val(File7_20.Charges & '') + Val(File7_20.charges2 & '')) as Purchase from file7_20 inner join file_20h on file7_20.doc_no = FILE7_20H.DOC_NO "
If IsDate(xDate2.Text) Then cField1 = cField1 & turnFound(cField1) & " Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField1 = cField1 & turnFound(cField1) & " FILE7_20H.Store = " & MyParn(xstore.BoundText)

cField2 = "Select Sum(Val(File7_20H.DISCOUNT & '')) as Field1 from file7_20H WHERE FILE7_20H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField2 = cField2 & turnFound(cField2) & " Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField2 = cField2 & turnFound(cField2) & " FILE7_20H.Store = " & MyParn(xstore.BoundText)
 
' „—œÊœ „‘ —Ì« 
cField3 = "Select Sum(Val(FILE7_30.Price & '') * Val(FILE7_30.Quant & '')) as Purchase from FILE7_30 inner join file_20h on FILE7_30.doc_no = FILE7_30H.DOC_NO WHERE FILE7_30H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField3 = cField3 & turnFound(cField3) & " FILE7_30H.Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField3 = cField3 & turnFound(cField3) & " FILE7_30H.Store = " & MyParn(xstore.BoundText)

cField4 = "Select Sum(Val(FILE7_30H.DISCOUNT & '')) as Field1 from FILE7_30H WHERE FILE7_30H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField4 = cField4 & turnFound(cPurchaseRetDiscount) & " FILE7_30H.Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField4 = cField4 & turnFound(cField4) & " FILE7_30H.Store = " & MyParn(xstore.BoundText)

' „»Ì⁄« 
cField5 = "Select Sum(Val(FILE6_20.Price & '') * Val(FILE6_20.Quant & '')) as Sales from FILE6_20 inner join file_20h on FILE6_20.doc_no = FILE6_20H.DOC_NO WHERE FILE6_20H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField5 = cField5 & turnFound(cField5) & " Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField5 = cField5 & turnFound(cField5) & " FILE7_30H.Store = " & MyParn(xstore.BoundText)

cField6 = "Select Sum(Val(FILE6_20H.DISCOUNT & '')) as Field1 from FILE6_20H WHERE FILE6_20H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField6 = cField6 & " FILE6_20H.Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField6 = cField6 & turnFound(cField6) & " FILE6_20H.Store = " & MyParn(xstore.BoundText)


' „—œÊœ „»Ì⁄« 
cField7 = "Select Sum(Val(FILE6_30.Price & '') * Val(FILE6_30.Quant & '')) as Salesret from FILE6_30 inner join file_30h on FILE6_30.doc_no = FILE6_30H.DOC_NO WHERE FILE6_30H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField7 = cField7 & " Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField7 = cField7 & turnFound(cField7) & " FILE6_30H.Store = " & MyParn(xstore.BoundText)

cField8 = "Select Sum(Val(FILE6_30H.DISCOUNT & '')) as Field1 from FILE6_30H WHERE FILE6_30H.DATE > " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cField8 = cField8 & " FILE6_30H.Date <= " & DateSq(xDate2.Text)
If xstore.BoundText <> "" Then cField8 = cField8 & turnFound(cField8) & " FILE6_30H.Store = " & MyParn(xstore.BoundText)

'  ÕÊÌ·«  „‰
If Trim(xstore.BoundText) <> "" Then
    cField9 = "Select Sum(Val(FILE1_60.QUANT & '') * Val(FILE1_60.Cost & '')) as Field from FILE1_60 inner join file_60h on FILE1_60.doc_no = FILE1_60H.DOC_NO WHERE FILE1_60H.DATE > " & DateSq(xDate1.Text) & _
                 "Store1 = " & MyParn(xstore.BoundText)
    If IsDate(xDate2.Text) Then cField9 = cField9 & " FILE1_60H.Date <= " & DateSq(xDate2.Text)
    If xstore.BoundText <> "" Then cField9 = cField9 & turnFound(cField9) & " FILE6_30H.Store = " & MyParn(xstore.BoundText)
End If

cString = "Select File8_52.DescA,Sum(Value) as SumofValue," & _
          " From (file8_50 Inner Join File8_51 on File8_50.Charge = File8_51.Code) left join file8_52 on file8_51.mainGroup = File8_52.code "

End Sub
