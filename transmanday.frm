VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form transmandayfrm 
   Caption         =   "ÊÍæíáÇÊ Çáíæã"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   7500
      Left            =   1395
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   8130
      _cx             =   14340
      _cy             =   13229
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   63.727959697733
      ZoomMode        =   4
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   5055
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9555
      _cx             =   16854
      _cy             =   8916
      _ConvInfo       =   1
      Appearance      =   0
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5085
      Width           =   3120
      Begin VB.CommandButton Command2 
         Caption         =   "ØÈÇÚÉ"
         Height          =   465
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÎÑæÌ"
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
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
Attribute VB_Name = "transmandayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sbox As String, sBoxName As String, sDate As String

Private Sub Command2_Click()
doprint
End Sub

Private Sub Form_Load()
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myLoad
End Sub
Private Sub myLoad()
cField1 = myiif2("NO1 = " & MyParn(Sbox), "[VALUE]") & " AS [ãÓÍæÈÇÊ]"
cField2 = myiif2("NO2 = " & MyParn(Sbox), "[VALUE]") & " AS [ÇíÏÇÚ]"
cField3 = myiif2("NO1 = " & MyParn(Sbox), "[VALUE]", "-1 * [VALUE]") & " AS " & "[ÕÇÝí ÇáãÓÍæÈÇÊ]"
cString = "SELECT FILE0_50.DESCA AS [ÓÍÈ ãä],FILE0_50_1.DESCA AS [ÇíÏÇÚ Ýí],FILE0_51.DESCA AS [ÇáÈíÇä]," & cField1 & "," & cField2 & "," & cField3 & _
          " FROM (FILE0_51 INNER JOIN FILE0_50 ON FILE0_51.NO1 = FILE0_50.CODE) INNER JOIN FILE0_50 AS FILE0_50_1 ON FILE0_51.NO2 = FILE0_50_1.CODE"
cString = cString & turn(cString) & "DATE = " & DateSq(sDate)
cString = cString & turn(cString) & "(NO1 = " & MyParn(Sbox) & " or NO2 = " & MyParn(Sbox) & ")"
data1.RecordSource = cString
data1.Refresh
Fixgrd
End Sub
Private Sub Fixgrd()
Dim nTotal1 As Single, nTotal2 As Single, nTotal3 As Single
grid1.ColWidth(0) = 1300
grid1.ColWidth(1) = 1300
grid1.ColWidth(2) = 3500
grid1.ColWidth(3) = 1000
grid1.ColWidth(4) = 1000
grid1.ColWidth(5) = 1000

For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
CalcTotals
End Sub
Private Sub CalcTotals()
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 3) = Myvalue(grid1.TextMatrix(i, 3), "FIXED")
    grid1.TextMatrix(i, 4) = Myvalue(grid1.TextMatrix(i, 4), "FIXED")
    grid1.TextMatrix(i, 5) = Myvalue(grid1.TextMatrix(i, 5), "FIXED")
    nTotal1 = nTotal1 + Val(grid1.TextMatrix(i, 3))
    nTotal2 = nTotal2 + Val(grid1.TextMatrix(i, 4))
    nTotal3 = nTotal3 + Val(grid1.TextMatrix(i, 5))
Next
grid1.AddItem ""
grid1.MergeCells = flexMergeFree
grid1.MergeRow(grid1.Rows - 1) = True
grid1.TextMatrix(grid1.Rows - 1, 0) = "ÇáÇÌãÇáí"
grid1.TextMatrix(grid1.Rows - 1, 1) = "ÇáÇÌãÇáí"
grid1.TextMatrix(grid1.Rows - 1, 2) = "ÇáÇÌãÇáí"
grid1.TextMatrix(grid1.Rows - 1, 3) = Myvalue(nTotal1, "fixed")
grid1.TextMatrix(grid1.Rows - 1, 4) = Myvalue(nTotal2, "fixed")
grid1.TextMatrix(grid1.Rows - 1, 5) = Myvalue(nTotal3, "fixed")
grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = vbYellow
End Sub
Private Sub doprint()
'vp.Visible = True
With vp
     vp = " "
    .Device = pDevice
    .StartDoc
    .fontsize = 8
    .TextColor = vbBlack
    .FontName = "Arial"
    .MarginLeft = 150
    .TextAlign = taCenterTop
    '.PenStyle = psTransparent
    .Paragraph = "iPlanet "
    .Paragraph = "ÊÝÕíáí ÊÍæíá ÎÒäÉ : " & sBoxName
    .Paragraph = "íæã : " & sDate
    .TextAlign = taRightTop
    .Paragraph = String(40, "=")
    .Paragraph = "ÊÇÑíÎ : " & Format(Date, "DD-MM-YYYY")
    .Paragraph = "æÞÊ : " & Time
    .Paragraph = String(40, "=")
    
    f = ">700|>2400|>1100;"
    H = ""
    For i = 1 To grid1.Rows - 2
        If Val(grid1.TextMatrix(i, 3)) <> 0 Then
            cRow = grid1.TextMatrix(i, 3) & "|" & _
            grid1.TextMatrix(i, 2) & "|" & ArbString(grid1.TextMatrix(i, 1) & "(Ó)") & ";"
            .AddTable f, H, cRow
        End If
        If Val(grid1.TextMatrix(i, 4)) <> 0 Then
            cRow = grid1.TextMatrix(i, 5) & "|" & _
            grid1.TextMatrix(i, 2) & "|" & ArbString(grid1.TextMatrix(i, 0) & "(Ç)") & ";"
            .AddTable f, H, cRow
        End If
    Next
    f = ">700|<3500;"
    H = ""
    cRow = grid1.TextMatrix(i, 5) & "|" & _
         "ÇáÇÌãÇáí" & ";"
    .AddTable f, H, cRow
    .EndDoc
End With
End Sub


