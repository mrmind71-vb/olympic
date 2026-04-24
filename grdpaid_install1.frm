VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form grdpaid_Installfrm1 
   Caption         =   "ÅÌãÇáí ÇíÕÇáÇÊ ÇáÇÚÖÇÁ"
   ClientHeight    =   10110
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   16785
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   16785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "äæÚ ÇáÓÏÇÏ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   270
      Width           =   5235
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Çáßá"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "äÞÏí"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   2227
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ÝæÑí"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   315
      Width           =   4965
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3735
         Picture         =   "grdpaid_install1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdpaid_install1.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grdpaid_install1.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   2430
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdpaid_install1.frx":7149
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdpaid_install1.frx":A147
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   9735
      Width           =   16785
      _ExtentX        =   29607
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   10395
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   6315
      Begin VB.TextBox xCode 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   1005
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
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
         TabIndex        =   1
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "ÑÞã ÇáÚÖæíÉ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   675
         Width           =   1050
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   3840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì ÊÇÑíÎ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2520
      Top             =   405
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
      Left            =   900
      Top             =   585
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
      Height          =   7530
      Left            =   90
      TabIndex        =   9
      Top             =   1080
      Width           =   16620
      _cx             =   29316
      _cy             =   13282
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   14
      Top             =   9585
      Visible         =   0   'False
      Width           =   16785
      _ExtentX        =   29607
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grdpaid_Installfrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim oSearchMember As New Search, bNoCheck As Boolean
Dim con As New ADODB.Connection
Dim aHeader()
Private Sub Check1_Click()
If Not bNoCheck Then
    myload
End If
bNoCheck = False
End Sub
Private Sub cmdExel_Click()
ToFileExel grid1, , , , 0.9
End Sub
Private Sub cmdPrint_Click()
Dim nRate As Double, i As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 5)
aRow(0) = AddFlag(aRow(0), "text", "ÇáÅÌãÇáí")
Dim nwidth As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = grid1.ColWidth(i) + nwidth
    End If
Next
nRate = 15000 / nwidth
Set PrintGrdNew.myForm = Me
'grid1.ColHidden(2) = True
PrintGrdNew.doprint grid1, nRate, -2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, True, 10, , aRow
'grid1.ColHidden(2) = False
PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set grdpaid1 = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub
Private Sub Form_Load()
Me.Top = 1000
Me.Left = 1000
openCon con

Set grid1.DataSource = data10
Fixgrd
grid1.ExplorerBar = flexExSortShow
bNoCheck = True
LoadText Me
If ValidNum(xCode.text) Then xCode_LostFocus
bNoCheck = False
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(5)
With grid1
If Not MYVALID Then Exit Sub
cString = "SELECT ''," & _
          " FILE6_30H.FORM_NO," & _
          " FILE6_30H.CODE," & _
          " FILE2_10.DESCA," & _
          " CONVERT(VARCHAR(10),FILE6_30H.DATE,111)," & _
          " COALESCE(SUM(FILE6_30.VALUE),0)," & _
          " COALESCE(SUM(FILE6_30.TAX),0)   +  COALESCE(SUM(FILE6_30.TAX_DIFF),0) + OTHER_VALUE_TAX," & _
          " FILE6_30H.CARD_VALUE," & _
          " COALESCE(SUM(FILE6_30.INTEREST),0) AS INTEREST," & _
          " FILE6_30H.OTHER," & _
          " SUM(FILE6_30.CHARGE)," & _
          " COALESCE(SUM(FILE6_30.TOTAL),0) + FILE6_30H.CARD_VALUE + FILE6_30H.OTHER + OTHER_VALUE_TAX " & _
          " FROM FILE6_30H " & _
          " LEFT JOIN FILE6_30 ON FILE6_30H.DOC_NO = FILE6_30.DOC_NO" & _
          " INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE"
cWhere = "(NOT FILE6_30H.FORM_NO IS NULL)"

If ValidNum(xCode.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE6_30H.CODE = " & xCode.text
    aHeader(0) = "ÇáÚÖæ : " & xCodeDesca.Caption
End If

If IsDate(xDate1.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE6_30H.DATE >= " & DateSq(xDate1.text)
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If
          
If IsDate(xDate2.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE6_30H.DATE <= " & DateSq(xDate2.text)
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

If Option1(1).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30h.[IsFawry] = 0"
    aHeader(2) = "ÓÏÇÏ " & Option1(1).Caption
ElseIf Option1(2).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_30h.[IsFawry] = 1"
    aHeader(2) = "ÓÏÇÏ " & Option1(2).Caption
End If

If cWhere <> "" Then cString = cString & " AND " & cWhere
cString = cString & " GROUP BY FILE6_30H.FORM_NO,FILE6_30H.CARD_VALUE,FILE6_30H.DATE,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.OTHER,OTHER_VALUE_TAX"
cString = cString & " ORDER BY FILE6_30H.DATE,FILE6_30H.FORM_NO"
Set data10.Recordset = myRecordSet(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ãÓáÓá"
    .TextMatrix(0, 1) = "ÑÞã ÇáÇíÕÇá"
    .TextMatrix(0, 2) = "ÑÞã ÇáÚÖæíÉ"
    .TextMatrix(0, 3) = "ÇÓã ÇáÚÖæ"
    .TextMatrix(0, 4) = "ÊÇÑíÎ ÇáãÓÊäÏ"
    .TextMatrix(0, 5) = "ÞíãÉ ÇáÞÓØ"
    .TextMatrix(0, 6) = "ÇáÖÑíÈÉ"
    .TextMatrix(0, 7) = "ÞíãÉ ÇáßÑæÊ"
    .TextMatrix(0, 8) = "ÇáÝÇÆÏÉ"
    .TextMatrix(0, 9) = "ÊÈÑÚ"
    .TextMatrix(0, 10) = "ãÕÇÑíÝ ÇÏÇÑíÉ"
    .TextMatrix(0, 10 + 1) = "ÇáÅÌãÇáí"
    
        
    .ColWidth(0) = 900
    .ColWidth(1) = 1000
    .ColWidth(2) = 1300
    .ColWidth(3) = 2500
    .ColWidth(4) = 1200
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1400
    .ColWidth(10) = 1300
    .ColWidth(10 + 1) = 1600
    
    .ColDataType(1) = flexDTLong
    .ColDataType(2) = flexDTLong
    .ColDataType(4) = flexDTDate
    .ColDataType(5) = flexDTDecimal
    .ColDataType(6) = flexDTDecimal
    .ColDataType(7) = flexDTDecimal
    .ColDataType(8) = flexDTDecimal
    .ColDataType(9) = flexDTDecimal
    .ColDataType(10) = flexDTDecimal
    .ColDataType(11) = flexDTDecimal

'    .ColHidden(0) = True

'    .ColDataType(2) = flexDTDouble
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble

    
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 5, "#0", &HC0FFC0, vbBlack, True, "  "
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0.00", &HC0FFC0, vbBlack, True, "  "
    Next
    
    FixSerial
               
    If .rows > 1 Then
        .TextMatrix(1, 1) = "ÇáÅÌãÇáì"
        For i = 6 To .Cols - 1
            .TextMatrix(1, i) = mRound(.TextMatrix(1, i))
        Next
    End If
    SBar1.Panels(1).text = IIf(grid1.rows > 2, "ÚÏÏ ÇáÓÌáÇÊ : " & grid1.rows - 2, "")
    End With
End Sub
Private Sub FixSerial()
Dim i As Long
For i = 2 To grid1.rows - 1
    grid1.TextMatrix(i, 0) = i - 1
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdbankfrm1 = Nothing
End Sub
Private Sub cmdPrintLand_Click()
Dim nRate As Double, i As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)
aRow(0) = AddFlag(aRow(0), "text", "ÇáÅÌãÇáí")
Dim nwidth As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = grid1.ColWidth(i) + nwidth
    End If
Next
nRate = 15000 / nwidth
Set PrintGrdNew.myForm = Me
PrintGrdNew.doprint grid1, nRate, -2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, True, 11, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
FixSerial
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub

Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xdesca.text, 1
End If
End Sub
Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Member_InLookupAll Me, oSearchMember
End If
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidNum(xCode.text) Then Exit Sub
'aRet = GetField("select DESCA from file1_10 where code = " & xCode.Text)
'If Not IsEmpty(aRet) Then xCodeDesca.Caption = retFlag(aRet, "DESCA") & ""
xCodeDesca.Caption = GetField("select DESCA from file2_10 where code = " & addvalue(xCode.text), con) & ""
End Sub
Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.rows > 1
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xbox_GotFocus()
myGotFocus xbox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Sub myProc()
xCode.text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
Unload oSearchMember
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub xShare_Click(Area As Integer)

End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xgroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
