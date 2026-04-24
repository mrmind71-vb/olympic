VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form vsman 
   Caption         =   "≈” ⁄Ê«÷ „‰ „Œ«“‰ «·‘—ﬂ…"
   ClientHeight    =   10095
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
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
      Height          =   375
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   990
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
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
      Height          =   375
      Left            =   990
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   990
      Width           =   915
   End
   Begin VB.CommandButton Cmd_Print 
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
      Height          =   375
      Left            =   1935
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   990
      Width           =   915
   End
   Begin VB.CommandButton CmdOk 
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
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   990
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   11835
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   7125
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   1455
      End
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   2925
         TabIndex        =   4
         Top             =   900
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Height          =   195
         Left            =   6060
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   6060
         TabIndex        =   6
         Top             =   195
         Width           =   675
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
         Left            =   6030
         TabIndex        =   5
         Top             =   585
         Width           =   735
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
      Height          =   8610
      Left            =   45
      TabIndex        =   0
      Top             =   1395
      Width           =   18915
      _cx             =   33364
      _cy             =   15187
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
Attribute VB_Name = "vsman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï ⁄œœ »Ê‰«  Ê ﬁÌ„… „»Ì⁄«  ·»«∆⁄Ì‰"
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.doprint Me.VsItem, 1, -2, cHead1, cHead2, , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Sub MyLoad()
Dim nPack As Double
Dim nBal2 As Double
With VsItem
.ExplorerBar = flexExSortShow
.Rows = 1
.SubtotalPosition = flexSTBelow

If IsDate(xDate1.Text) Then cWhere = cWhere & turnFound(cWhere, " and ") & " Date >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cWhere = cWhere & turnFound(cWhere, " and ") & " Date <= " & DateSql(xDate2.Text)
If xStore.BoundText <> "" Then cWhere = cWhere & turnFound(cWhere, " and ") & " STORE = " & MyParn(xStore.BoundText)

cField1 = "(Select Sum(1) From file6_20H " & turnFound(cWhere) & cWhere & " ) as CountDoc"
cField2 = myiif("SALESDTL.Flag = 0", "Total") & " AS TTSAL"
cField3 = myiif("SALESDTL.Flag = 1", "Total") & " AS TTRET"
        
cString = "SELECT MAN , FILE6_25.DESCA " & _
          "," & cField1 & _
           "," & cField2 & _
           "," & cField3 & " FROM SalesDtl Left join file6_25 on salesDtl.Man = file6_25.code  "
cString = cString & turnFound(cWhere) & cWhere
cString = cString & " GROUP BY MAN , FILE6_25.DESCA"
Dim datatable As New ADODB.Recordset
datatable.Open cString, CON, adOpenStatic, adLockReadOnly, adCmdText
Do Until datatable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = datatable!MAN & ""
    .TextMatrix(.Rows - 1, 1) = datatable!desca & ""
    .TextMatrix(.Rows - 1, 2) = Format(datatable!countdoc, "#0")
    .TextMatrix(.Rows - 1, 3) = Format(datatable!ttsal, "#0.00")
    .TextMatrix(.Rows - 1, 4) = Format(Val(datatable!ttret & "") * -1, "#0.00")
    .TextMatrix(.Rows - 1, 5) = Format(Val(.TextMatrix(.Rows - 1, 3)) - Val(.TextMatrix(.Rows - 1, 4)), "#0.00")
    datatable.MoveNext
Loop
.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "#0", , vbRed, , " "
.Subtotal flexSTSum, -1, 3, "#0.00", , vbRed, , " "
.Subtotal flexSTSum, -1, 4, "#0.00", , vbRed, , " "
.Subtotal flexSTSum, -1, 5, "#0.00", , vbRed, , " "
End With
datatable.Close
Set datatable = Nothing
End Sub

Private Sub CmdOk_Click()
MyLoad
End Sub

Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

With VsItem
.Cols = 6
.Rows = 1
.RowHeight(0) = 1000
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·»«∆⁄"
.TextMatrix(0, 2) = "⁄œœ »Ê‰« "
.TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 4) = "ﬁÌ„… „— Ã⁄"
.TextMatrix(0, 5) = "«·’«›Ï"

.ColWidth(0) = 600
.ColWidth(1) = 2000
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500

.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
End With
End Sub
