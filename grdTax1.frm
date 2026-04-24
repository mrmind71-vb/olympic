VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdTax1 
   Caption         =   "ÊÝÕíáí ãÓÊÍÞÇÊ ÖÑíÈÉ ÞíãÉ ãÖÇÝÉ "
   ClientHeight    =   10110
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   20250
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
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   630
      Width           =   4965
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3735
         Picture         =   "grdTax1.frx":0000
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
         Picture         =   "grdTax1.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grdTax1.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmdPrintLand 
         Height          =   555
         Left            =   2430
         TabIndex        =   14
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
         Picture         =   "grdTax1.frx":7149
         Caption         =   "ØÈÇÚÉ"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdTax1.frx":A147
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   9735
      Width           =   20250
      _ExtentX        =   35719
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
      Height          =   1410
      Left            =   12375
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -45
      Width           =   7800
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
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
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
         Left            =   4815
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì"
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   300
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   2715
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2790
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
      Left            =   1170
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
      Height          =   7575
      Left            =   90
      TabIndex        =   8
      Top             =   1395
      Width           =   20085
      _cx             =   35428
      _cy             =   13361
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Cols            =   9
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
      TabIndex        =   13
      Top             =   9585
      Visible         =   0   'False
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grdTax1"
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
Dim myHeader(2) As Variant
myHeader(0) = Me.Caption
myHeader(1) = retHeader(aHeader, 0, 1)
myHeader(2) = retHeader(aHeader, 1, 3)

ToFileExel2 grid1, , , , , , , , aSplit, , , Me, myHeader
End Sub
Private Sub CmdPrint_Click()
Dim nRate As Double, I As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 3)
aRow(0) = AddFlag(aRow(0), "text", "ÇáÅÌãÇáí")
Dim nWidth As Double
For I = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(I) Then
        nWidth = grid1.ColWidth(I) + nWidth
    End If
Next
nRate = grdRate(grid1, 15500)
Set PrintGrdNew.myForm = Me
grid1.ColHidden(1) = True
grid1.ColHidden(4) = True
grid1.ColHidden(5) = True

Me.MousePointer = 11
PrintGrdNew.doprint grid1, nRate, -2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, False, 14, , aRow
grid1.ColHidden(1) = False
grid1.ColHidden(4) = False
grid1.ColHidden(5) = False
Me.MousePointer = 0
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
Me.MousePointer = 11
myload
Me.MousePointer = 0
End Sub
Private Sub Form_Load()
openCon con

Set grid1.DataSource = data10
grid1.ExplorerBar = flexExSortShow
Fixgrd
bNoCheck = True
LoadText Me
bNoCheck = False
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(5)
With grid1
If Not MYVALID Then Exit Sub

cString = "SELECT '', FILE1_10.CODE AS [ÑÞã ÇáÚÖæíÉ], FILE1_10.DESCA AS [ÇáÇÓã],CONVERT(VARCHAR(10), FILE1_60.DATE,111) AS [ÇáÊÇÑíÎ]," & _
          " FILE1_60.VALUE AS [ÅÌãÇáí ÇáÞÓØ], FILE1_60.TAX AS [ÇáÖÑíÈÉ], FILE1_60.MONTHS_LATE AS [ÔåæÑ ÇáÛÑÇãÉ], FILE1_60.TAX_LATE AS [ÇáÛÑÇãÉ], FILE1_60.TOTAL AS [ÅÌãÇáí ÇáãÓÊÍÞ]" & _
          " FROM FILE1_60 INNER JOIN FILE1_10 ON FILE1_60.MEMBER = FILE1_10.CODE"

If ValidNum(xCode.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_60.[MEMBER]  = " & addvalue(xCode.Text)
End If

If IsDate(xDate1.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_60.[DATE] >= " & DateSq(xDate1.Text, True)
    aHeader(1) = BetweenString(xDate1.Text, xDate2.Text)
End If
          
If IsDate(xDate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_60.[DATE] <= " & DateSq(xDate2.Text, True)
    aHeader(1) = BetweenString(xDate1.Text, xDate2.Text)
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

cString = cString & " ORDER BY FILE1_60.DATE,FILE1_60.MEMBER"
Set data10.Recordset = myRecordSet(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd(Optional bPrint As Boolean = False)
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ãÓáÓá"
           
    .ColWidth(0) = 800
    .ColWidth(1) = 900
    .ColWidth(2) = 2600
    .ColWidth(3) = 1300
    .ColWidth(4) = 1300
    .ColWidth(5) = 1300
    .ColWidth(6) = 800
    .ColWidth(7) = 1200
    .ColWidth(8) = 1400
    
    .TextMatrix(0, 1) = "ÑÞã ÇáÚÖæíÉ"
    .TextMatrix(0, 2) = "ÇáÇÓã"
    .TextMatrix(0, 3) = "ÇáÊÇÑíÎ"
    .TextMatrix(0, 4) = "ÅÌãÇáí ÇáÞÓØ"
    .TextMatrix(0, 5) = "ÇáÖÑíÈÉ"
    .TextMatrix(0, 6) = "ÔåæÑ ÇáÛÑÇãÉ"
    .TextMatrix(0, 7) = "ÇáÛÑÇãÉ"
    .TextMatrix(0, 8) = "ÅÌãÇáí ÇáãÓÊÍÞ"
    .ColDataType(3) = flexDTDate
         
    
    For I = 1 To grid1.rows - 1
        .TextMatrix(I, 0) = I
    Next
    
    .SubtotalPosition = flexSTAbove
    
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
        If I >= 4 And I <> 6 Then
            .Subtotal flexSTSum, -1, I, IIf(bPrint Or I = 4, "#0", "#0.00"), &HC0FFC0, vbBlack, True, "  "
            .ColDataType(I) = flexDTDouble
        End If
    Next
               
    If .rows > 1 Then
        .TextMatrix(1, 1) = "ÇáÅÌãÇáì"
    End If
    SBar1.Panels(1).Text = IIf(grid1.rows > 2, "ÚÏÏ ÇáÓÌáÇÊ : " & grid1.rows - 2, "")
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdbankfrm1 = Nothing
End Sub
Private Sub cmdPrintLand_Click()
Dim nRate As Double, I As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)
aRow(0) = AddFlag(aRow(0), "text", "ÇáÅÌãÇáí")
Fixgrd True
nRate = grdRate(grid1)
Set PrintGrdNew.myForm = Me
PrintGrdNew.doprint grid1, nRate, 2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, False, 11, , aRow
PrintGrdNew.Show 1
Fixgrd
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub

Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDesca.Text, 1
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearchMember
End If
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidNum(xCode.Text) Then Exit Sub
'aRet = GetField("select DESCA from file1_10 where code = " & xCode.Text)
'If Not IsEmpty(aRet) Then xCodeDesca.Caption = retFlag(aRet, "DESCA") & ""
xCodeDesca.Caption = GetField("select DESCA from file1_10 where code = " & addvalue(xCode.Text), con) & ""
End Sub

Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
'cmdPrint.Enabled = grid1.rows > 1
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xcode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xcode2
End Sub
Sub myProc()
xCode.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
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

Private Sub xFile_name_Change()

End Sub
