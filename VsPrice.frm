VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form VsPrice 
   Caption         =   "„ «»⁄… «”⁄«— «·»Ì⁄"
   ClientHeight    =   9195
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12885
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
   ScaleHeight     =   9195
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   8775
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   150
         Width           =   1455
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   510
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   3450
         TabIndex        =   10
         Top             =   870
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   3450
         TabIndex        =   11
         Top             =   510
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   3450
         TabIndex        =   12
         Top             =   150
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7005
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   555
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7005
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   1230
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
         Index           =   0
         Left            =   1605
         TabIndex        =   8
         Top             =   555
         Width           =   735
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
         Left            =   1635
         TabIndex        =   7
         Top             =   165
         Width           =   675
      End
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
      Left            =   2925
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1035
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
      Left            =   1980
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1035
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
      Left            =   1035
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1035
      Width           =   915
   End
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
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1035
      Width           =   915
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7LCtl.VSFlexGrid grid1 
      Height          =   7470
      Left            =   90
      TabIndex        =   9
      Top             =   1440
      Width           =   12720
      _cx             =   22437
      _cy             =   13176
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   60
      Top             =   300
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   525
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Top             =   750
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
Attribute VB_Name = "VsPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cString2 As String
Dim cFilter1 As String
Dim cFilter As String
Dim nInd As Byte
Private Sub Cmd_Print_Click()
Dim temptable As New ADODB.Recordset
contemp.Execute "Delete * From Temp"
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
Dim aHeader(1)
With grid1
    For I = 2 To .Rows - 1
        temptable.AddNew
        temptable!str1 = TurnValue(.TextMatrix(I, 1), "", Null)
        temptable!str2 = TurnValue(.TextMatrix(I, 2), "", Null)
        
        temptable!val1 = Val(.TextMatrix(I, 3))
        temptable!val2 = Val(.TextMatrix(I, 4))
        temptable!val3 = Val(.TextMatrix(I, 5))
        temptable!val4 = Val(.TextMatrix(I, 6))
        temptable!val5 = Val(.TextMatrix(I, 7))
        temptable!Val6 = Val(.TextMatrix(I, 8))
        temptable!Val7 = Val(.TextMatrix(I, 9))

        temptable!str21 = "„ «»⁄… «·„»Ì⁄«  & «”⁄«— «·»Ì⁄  "
        temptable!str22 = retHeader(aHeader, 0, 1)
        If IsDate(xDate1.Text) Or IsDate(xDate2.Text) Then
            aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
        End If
        temptable.Update
    Next
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = PublicPath & "\Reports\vsprice.rpt"
main.Report1.DataFiles(0) = cPathTemp
main.Report1.Action = 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdOk_Click()
'If IsDate(xDate1.Text) And IsDate(xDate2.Text) Then
   fillgrd
End Sub
Sub fillgrd()
Dim nSubTotal As Double
Dim nTotal As Double
Dim nT As Double
Dim cSubGroup As String, CGROUP As String
Dim invTable As New ADODB.Recordset
Me.MousePointer = 11

grid1.Sort = flexSortNone
xGroup.Enabled = True
nSubTotal = 0
nTotal = 0
nT = 0
cFilter = ""
cString = "SELECT Min(FILE6_20.PRICE) AS MinPRICE, Max(FILE6_20.PRICE) AS MaxPRICE, Sum(FILE6_20.TOTAL) AS  " & _
          " SumTOTAL, Sum(FILE6_20.QUANT) AS SumQUANT, FILE1_10.ITEM, FILE1_10.DESCA  , FILE1_10.PRICE  " & _
          " FROM ((FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO ) inner join file1_50 on file1_10.[GROUP] = file1_50.code "
    If xGroup.BoundText <> "" Then cStrAll = cStrAll & " AND file1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStrAll = cStrAll & " AND file1_50.[Group]  = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStrAll = cStrAll & " AND [Section] = " & xSection.BoundText

If xGroup.BoundText <> "" Then
    cString = cString & turnFound2(cString) & " File1_10.[GROUP] = " & MyParn(xGroup.BoundText)
End If

If IsDate(xDate1.Text) Then
    cString = cString & turnFound2(cString) & " file6_20H.Date >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound2(cString) & " file6_20H.Date <= " & DateSq(xDate2.Text)
End If

cString = cString & " GROUP BY FILE1_10.ITEM, FILE1_10.DESCA  , FILE1_10.PRICE  "

invTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
.FixedRows = 1
.FixedCols = 1
.Rows = 1

Do Until invTable.EOF
   .AddItem ""
   .TextMatrix(.Rows - 1, 1) = TurnValue(invTable!Item, Null, "")
   .TextMatrix(.Rows - 1, 2) = TurnValue(invTable!Desca, Null, "")
   .TextMatrix(.Rows - 1, 3) = Format(invTable!SUMQUANT, "##0.00")
   .TextMatrix(.Rows - 1, 4) = Format(invTable!Sumtotal, "##0.000")
   .TextMatrix(.Rows - 1, 5) = TurnValue(invTable!price, Null, "")
   .TextMatrix(.Rows - 1, 6) = Format(invTable!MAXPRICE, "##0.00")
   .TextMatrix(.Rows - 1, 7) = Format(invTable!MINPRICE, "##0.00")
   .TextMatrix(.Rows - 1, 8) = Format(Val(.TextMatrix(.Rows - 1, 5)) - Val(.TextMatrix(.Rows - 1, 7)), "##0.00")
    If Val(.TextMatrix(.Rows - 1, 8)) > 0 And Val(.TextMatrix(.Rows - 1, 6)) <> 0 Then .TextMatrix(.Rows - 1, 9) = Format(Val(.TextMatrix(.Rows - 1, 8)) * 100 / Val(.TextMatrix(.Rows - 1, 6)), "##0.00")
   invTable.MoveNext
Loop
'.Sort = 1
Me.MousePointer = 1
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 3, "##0.00", , RGB(255, 0, 0), True

.Subtotal flexSTSum, -1, 4, "##0.00", , RGB(255, 0, 0), True
.SubtotalPosition = flexSTAbove
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlue
End With
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE1_50"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

xDate1.Text = ""
xDate2.Text = ""
'Me.Picture = LoadPicture(App.Path & "\mainback.jpg")
With grid1
.ExplorerBar = flexExSortShow
.OutlineBar = flexOutlineBarNone
.Cols = 10
.Rows = 1
.TextMatrix(0, 1) = "ﬂÊœ"
.TextMatrix(0, 2) = "’‰›"
.TextMatrix(0, 3) = "Ã. „»Ì⁄« "
.TextMatrix(0, 4) = "Ã. ﬁÌ„…"
.TextMatrix(0, 5) = "”⁄— „” Â·ﬂ"
.TextMatrix(0, 6) = "√⁄·Ï ”⁄—"
.TextMatrix(0, 7) = "√ﬁ· ”⁄—"
.TextMatrix(0, 8) = "«·›—ﬁ"
.TextMatrix(0, 9) = "‰”»…"

.RowHeight(0) = 600
.WordWrap = True
.ColWidth(0) = 400
.ColWidth(1) = 1000
.ColWidth(2) = 2500
.ColWidth(3) = 800
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100

.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.FixedCols = 1
End With
End Sub
Private Sub grid1_DblClick()
    R_Sal_Item grid1.TextMatrix(grid1.Row, 1), grid1.TextMatrix(grid1.Row, 2)
End Sub
Sub R_Sal_Item(cItem, cDesc)
Dim aHeader(1)
Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset
contemp.Execute "Delete * From Temp"
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = " SELECT FILE3_10.DESCA, FILE6_20H.DATE, FILE6_20.QUANT, VAL(FILE6_20.QUANT & '') * VAL(FILE6_20.PRICE & '') as Total,  " & _
          " FILE6_20h.DOC_NO, FILE6_20.PRICE " & _
          " FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) INNER JOIN FILE3_10 ON FILE6_20H.CODE =  FILE3_10.CODE  " & _
          " where file6_20.item = " & MyParn(cItem)

If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "file6_20H.Date >= " & DateSq(xDate1.Text)
    aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
End If

If xDate2.Text <> "" Then
    cString = cString & turnFound(cString) & "file6_20H.Date >= " & DateSq(xDate1.Text)
    aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
End If
cString = cString & " ORDER BY  file6_20H.Date "

With sourcetable
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do While Not .EOF
    temptable.AddNew
    temptable!str2 = !doc_no
    temptable!str3 = !Desca
    temptable!Date1 = !Date
    temptable!val1 = !Quant
    temptable!val2 = !price
    temptable!val3 = !TOTAL
    
    temptable!str8 = "„»Ì⁄«  " & cItem & " ==> " & cDesc
    temptable!str9 = TurnValue(retHeader(aHeader, 0, 1))
'    temptable.STR19 = firsttitle
'    temptable.STR20 = Secondtitle
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.Reset
main.Report1.ReportFileName = PublicPath & "\Reports\SalItem.rpt"
main.Report1.DataFiles(0) = cPathTemp
main.Report1.Formulas(0) = "COL3 = " & MyParn("«·⁄„Ì·")
main.Report1.WindowState = crptMaximized
main.Report1.Action = 1
End Sub
