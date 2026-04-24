VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VsSal 
   Caption         =   "≈” ⁄Ê«÷ „‰ „Œ«“‰ «·‘—ﬂ…"
   ClientHeight    =   6480
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   375
      Left            =   8190
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   1185
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "⁄—÷"
      Height          =   375
      Left            =   8190
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   675
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   9405
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   9600
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   4995
         TabIndex        =   4
         Top             =   945
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   4995
         TabIndex        =   11
         Top             =   585
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   495
         TabIndex        =   13
         Top             =   540
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ „Œ“‰ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3690
         TabIndex        =   12
         Top             =   585
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„»Ì⁄«  „‰  «—ÌŒ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   8130
         TabIndex        =   8
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3645
         TabIndex        =   7
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄…"
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
         Left            =   8130
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   990
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Õ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8145
         TabIndex        =   5
         Top             =   675
         Width           =   495
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   150
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   9135
      Left            =   90
      TabIndex        =   0
      Top             =   1485
      Width           =   18915
      _cx             =   33364
      _cy             =   16113
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
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   900
      Top             =   315
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
   Begin MSAdodcLib.Adodc data2 
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
Attribute VB_Name = "VsSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datatable As New ADODB.Recordset
Dim Bal1Table As New ADODB.Recordset
Dim Bal2Table As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï √—’œ… „»Ì⁄«  „‘ —Ì«  ··√’‰«›  "
    cHead2 = " „‰  «—ÌŒ " & Format(xdate1.Text, "YYYY-MM-DD") & " ≈·Ï  «—ÌŒ " & Format(XDATE2.Text, "YYYY-MM-DD")
    
    Load PrintGrd
    PrintGrd.doprint Me.VsItem, 1, -2, cHead1, cHead2, , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdOk_Click()
If xStore.BoundText = "" Then
    MsgBox " ÕœÌœ «·„Õ·"
    Exit Sub
End If

If xStore2.BoundText = "" Then
    MsgBox " ÕœÌœ «·„Œ“‰"
    Exit Sub
End If

cField3 = myiif("FILE1_11!TYPE = '6' ", "[OUT]") & " as T_Sal ,"
cField4 = myiif("FILE1_11!TYPE = '6' ", "[OUT] * FILE1_11.PRICE ") & " as TV_Sal "

cString = "SELECT FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.COST1 , FILE1_10.COST2 , FILE1_10.PRICE , FILE1_10.COST4 , FILE1_10.[GROUP] ,  FILE1_10.R1 , FILE1_10.R2  , " & _
          cField3 & cField4 & _
          " FROM FILE1_11 LEFT JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"

If xStore.BoundText <> "" Then cString = cString & turnFound(cString) & " FILE1_11.STORE = " & MyParn(xStore.BoundText)
If xGroup.Text <> "" Then cString = cString & turnFound(cString) & "  FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
If IsDate(xdate1.Text) Then cString = cString & turnFound(cString) & " file1_11.DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cString = cString & turnFound(cString) & " file1_11.DATE <= " & DateSq(XDATE2.Text)
cString = cString & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PACK , FILE1_10.COST1, FILE1_10.COST2, FILE1_10.PRICE, FILE1_10.COST4, FILE1_10.[GROUP], R1 , R2 "

datatable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

cString = "SELECT FILE1_11.ITEM , SUM(FILE1_11.[IN]) AS TIN , SUM(FILE1_11.OUT) AS TOUT   " & _
          " FROM FILE1_11"

If xStore.Text <> "" Then cString = cString & turnFound(cString) & " FILE1_11.STORE= " & MyParn(xStore.BoundText)
cString = cString & " GROUP BY FILE1_11.ITEM "

Bal1Table.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

cString = "SELECT FILE1_11.ITEM , SUM(FILE1_11.[IN]) AS TIN , SUM(FILE1_11.OUT) AS TOUT   " & _
          " FROM FILE1_11"

If xStore2.Text <> "" Then cString = cString & turnFound(cString) & " FILE1_11.STORE = " & MyParn(xStore2.BoundText)
cString = cString & " GROUP BY FILE1_11.ITEM "

Bal2Table.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Me.MousePointer = 11

If Not (datatable.EOF And datatable.BOF) Then
    fillgrd
Else
    VsItem.Rows = 1
End If
Me.MousePointer = 0
End Sub
Sub fillgrd()
Dim nPack As Double
Dim nBal2 As Double
With VsItem
.FixedRows = 1
.FrozenCols = 4
.ExplorerBar = flexExSortShow
.Rows = 1
datatable.MoveFirst
.SubtotalPosition = flexSTBelow
Do Until datatable.EOF
    nBal2 = 0
    Bal2Table.Find " ITEM = " & MyParn(datatable!Item), , adSearchForward, adBookmarkFirst
    If Not Bal2Table.EOF Then nBal2 = TurnValue(Bal2Table!TIN, Null, 0) - TurnValue(Bal2Table!TOUT, Null, 0)
    If nBal2 > 0 Then
        If Val(datatable!R1 & "") > 0 Then
             .AddItem ""
            .TextMatrix(.Rows - 1, 0) = datatable!Item
            .TextMatrix(.Rows - 1, 1) = datatable!Desca
            .TextMatrix(.Rows - 1, 4) = datatable!price & ""
            .TextMatrix(.Rows - 1, 6) = TurnValue(datatable!R1, Null, 0)
            .TextMatrix(.Rows - 1, 7) = TurnValue(datatable!R2, Null, 0)
             
             Bal1Table.Find " ITEM = " & MyParn(datatable!Item), , adSearchForward, adBookmarkFirst
             If Not Bal1Table.EOF Then .TextMatrix(.Rows - 1, 5) = TurnValue(Bal1Table!TIN, Null, 0) - TurnValue(Bal1Table!TOUT, Null, 0)
             
             .TextMatrix(.Rows - 1, 9) = nBal2
             
             .TextMatrix(.Rows - 1, 2) = Format(TurnValue(datatable!T_SAL, NUL, 0), "#0")
             .TextMatrix(.Rows - 1, 3) = Format(TurnValue(datatable!TV_SAL, NUL, 0), "#0.00")
             If Val(.TextMatrix(.Rows - 1, 5)) < datatable!R2 Then
                 .TextMatrix(.Rows - 1, 8) = Val(.TextMatrix(.Rows - 1, 6)) - Val(.TextMatrix(.Rows - 1, 5))
                 If Val(.TextMatrix(.Rows - 1, 8)) > nBal2 Then .TextMatrix(.Rows - 1, 8) = nBal2
             End If
        End If
    End If
    datatable.MoveNext
Loop
datatable.Close
Bal1Table.Close
Bal2Table.Close

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "#0", , vbRed, , " "
.Subtotal flexSTSum, -1, 3, "#0", , vbRed, , " "
.Subtotal flexSTSum, -1, 5, "#0", , vbRed, , " "
.Subtotal flexSTSum, -1, 8, "#0", , vbRed, , " "
End With
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Set ItemTable = mydb.OpenRecordset("select * from FILE1_10")
Set GrTable = mydb.OpenRecordset("select * from FILE1_50")

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE1_50"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = DATA2
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE0_40"
Set xStore2.RowSource = DATA2
xStore2.ListField = "Desca"
xStore.BoundColumn = "Code"

With VsItem
.Cols = 10
.Rows = 1
.RowHeight(0) = 1000
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·’‰›"
.TextMatrix(0, 2) = "„»Ì⁄« "
.TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 4) = "”⁄— «·„” Â·ﬂ"
.TextMatrix(0, 5) = "«·—’Ìœ " & xStore.Text
.TextMatrix(0, 6) = "«·—›"
.TextMatrix(0, 7) = "≈⁄«œ… «·ÿ·»"
.TextMatrix(0, 8) = "ﬂ„Ì… «·≈” ⁄Ê«÷"
.TextMatrix(0, 9) = "«·—’Ìœ " & xStore2.Text

.ColWidth(0) = 1200
.ColWidth(1) = 3500
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 800
.ColWidth(5) = 800
.ColWidth(6) = 800
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 800

.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
End With
End Sub
Function myiif(cCondition, cField)
If cCondition = "" Then
    myiif = "Sum(" & cField & ")"
Else
    myiif = "Sum(iif(" & cCondition & "," & _
         cField & "," & "0" & "))"
End If
End Function


