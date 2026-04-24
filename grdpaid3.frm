VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdpaid3 
   Caption         =   "≈Ã„«·Ì ”œ«œ ÌÊ„Ì"
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
      Caption         =   "‰Ê⁄ «·”œ«œ"
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
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   675
      Width           =   5235
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
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
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "‰ﬁœÌ"
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
         TabIndex        =   16
         Top             =   360
         Width           =   825
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "›Ê—Ì"
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
         TabIndex        =   15
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5445
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   4920
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "grdpaid3.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   2430
         Picture         =   "grdpaid3.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdpaid3.frx":491C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "grdpaid3.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
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
      Height          =   1455
      Left            =   10395
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6315
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
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   2205
         TabIndex        =   11
         Top             =   990
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label19 
         Caption         =   "„Ã„Ê⁄… «·»‰œ"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   315
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1890
      Top             =   45
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
      Left            =   45
      Top             =   135
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
      Height          =   8205
      Left            =   180
      TabIndex        =   5
      Top             =   1485
      Width           =   16620
      _cx             =   29316
      _cy             =   14473
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
      Cols            =   8
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
End
Attribute VB_Name = "grdpaid3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim oSearchMember As New Search3
Dim con As New ADODB.Connection
Dim aHeader()
Private Sub cmdExel_Click()
ToFileExel grid1
End Sub
Private Sub CmdPrint_Click()
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 1)
Me.MousePointer = 11
PrintGrdNew.doprint grid1, 0.9, -2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, False, 10, , aRow
Me.MousePointer = 0
PrintGrdNew.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
Me.MousePointer = 11
myload
Me.MousePointer = 0
End Sub
Private Sub Form_Load()
Me.Top = 1000
Me.Left = 1000
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM FILE6_10G ORDER BY CODE", con)
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set grid1.DataSource = data10
Fixgrd
LoadText Me
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(2)
With grid1

cString = "SELECT CONVERT(VARCHAR(10),FILE6_20H.DATE,111),SUM(FILE6_20.QUANT),SUM(FILE6_20.TOTAL_ITEM),SUM(FILE6_20.DISCOUNT),SUM(FILE6_20.TOTAL_DISCOUNT),SUM(FILE6_20.TAX),SUM(FILE6_20.TOTAL)" & _
          " FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM  INNER JOIN FILE6_10G ON FILE6_10.[GROUP] = FILE6_10G.CODE" & _
          " WHERE (NOT FILE6_20H.FORM_NO IS NULL)"

If IsDate(xDate1.Text) Then
    cWhere = cString & turn(cWhere, " AND ") & "FILE6_20H.DATE >= " & DateSq(xDate1.Text)
    aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
End If
          
If IsDate(xDate2.Text) Then
    cWhere = cString & turn(cWhere, " AND ") & "FILE6_20H.DATE <= " & DateSq(xDate2.Text)
    aHeader(0) = BetweenString(xDate1.Text, xDate2.Text)
End If

If xGroup.MatchedWithList Then
    cWhere = cString & turn(cWhere, " AND ") & "FILE6_10.[GROUP] = " & addvalue(xGroup.BoundText)
    aHeader(1) = "«·„Ã„Ê⁄… : " & xGroup.Text
End If

If Option1(1).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20h.[IsFawry] = 0"
    aHeader(2) = "”œ«œ " & Option1(1).Caption
ElseIf Option1(2).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20h.[IsFawry] = 1"
    aHeader(2) = "”œ«œ " & Option1(2).Caption
End If

If cWhere <> "" Then cString = cString & " AND " & cWhere

cString = cString & " GROUP BY FILE6_20H.DATE"
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
    
    .TextMatrix(0, 0) = "«· «—ÌŒ"
    .TextMatrix(0, 1) = "«·≈Ã„«·Ì"
    .TextMatrix(0, 2) = "«·Œ’„"
    .TextMatrix(0, 3) = "’«›Ì  «·ﬁÌ„…"
    .TextMatrix(0, 4) = "«Ã„«·Ì «·÷—Ì»…"
    .TextMatrix(0, 5) = "«·≈Ã„«·Ì"
        
    .ColWidth(0) = 1400
    .ColWidth(1) = 2500
    .ColWidth(2) = 1400
    .ColWidth(3) = 1800
    .ColWidth(4) = 1500
    .ColWidth(5) = 1800
    
    .ColHidden(0) = True
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
           
    If .rows > 1 Then
        .SubtotalPosition = flexSTAbove
        For i = 2 To .Cols - 1
            .Subtotal flexSTSum, -1, i, "#.##", &HC0FFC0, vbBlack, True, "  "
        Next
        For i = 2 To .Cols - 1
            .TextMatrix(1, i) = mRound(.TextMatrix(1, i), 2)
        Next
        For i = 0 To 1
            .TextMatrix(1, i) = "«·≈Ã„«·Ï"
        Next
        .MergeCells = flexMergeFree
        .MergeRow(1) = True
    End If
    SBar1.Panels(1).Text = IIf(grid1.rows > 2, "⁄œœ «·”Ã·«  : " & grid1.rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdpaid3 = Nothing
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub
Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearchMember
End If
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidInt(xCode.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode.Text)
If Not IsEmpty(aRet) Then xCodeDesca.Caption = retFlag(aRet, "DESCA") & ""
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
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
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
'xCode.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
'xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
'Unload oSearchMember
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xDate2.Text) Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function

