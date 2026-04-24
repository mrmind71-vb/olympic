VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form showfrm3 
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Show3.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6165
      UseMaskColor    =   -1  'True
      Width           =   1680
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4200
      Top             =   5100
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6045
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   12705
      _cx             =   22410
      _cy             =   10663
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
End
Attribute VB_Name = "showfrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sHead1 As String, sHead2 As String, sHead3 As String, sWhere As String, sCol As String, sTotal As String
Public cWhere As String, nCounter As Long
Dim nSpan As Long, aSpan As Variant
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub CmdPrint_Click()
    PrintGrd.doprint grid1, 0.95, -2, sHead1, sHead2, sHead3, False, , 10, , Array(1), Array(1, 2)
    PrintGrd.Show 1
End Sub

Private Sub CmdExit_Click()
Unload Me
Set showfrm1 = Nothing
End Sub

Private Sub Form_Load()
Dim cString As String
Dim nWidth As Long
Me.Caption = sHead1
Set grid1.DataSource = data1
data1.ConnectionString = strCon
cString = "SELECT TOP 100 FILE8_50H.DOC_NO AS [—ﬁ„ «·„” ‰œ], FILE8_50.DESCA AS [«·»Ì«‰], FILE8_50.[VALUE] AS [«·ﬁÌ„…], COUNTER AS [«·⁄œ«œ],'' as [ﬂÌ·Ê„ — „‰ﬁ÷Ï]" & _
          "  FROM FILE8_50H INNER JOIN  FILE8_50 ON FILE8_50H.DOC_NO = FILE8_50.DOC_NO"
cString = cString & turn(cWhere) & cWhere
cString = cString & " ORDER BY FILE8_50.COUNTER DESC,FILE8_50H.DATE DESC"
data1.RecordSource = cString
data1.Refresh
FixGrd
With grid1
For i = 0 To grid1.Cols - 1
    nWidth = nWidth + Val(.ColWidth(i))
Next
grid1.Width = nWidth + (grid1.Cols * 50)
Me.Width = grid1.Left + grid1.Width + 200
End With
End Sub
Sub FixGrd()
Dim aLocal As Variant
With grid1

.RowHeight(0) = 800
.WordWrap = True

.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 1000
.ColWidth(3) = 1500
.ColWidth(4) = 1000
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, .Cols - 1) = nCounter - Val(grid1.TextMatrix(i, 3))
Next
End With
End Sub

