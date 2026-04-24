VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdmember1 
   Caption         =   "»Ì«‰«  « ’«· «·«⁄÷«¡"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11400
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
   ScaleHeight     =   10365
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   630
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grdMember1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "grdMember1.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdMember1.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "grdMember1.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   8415
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   10365
      Begin VB.CheckBox xDrop 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ…"
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
         Height          =   300
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1035
         Width           =   2040
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
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   540
         Width           =   1680
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   5805
         TabIndex        =   0
         Top             =   180
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   585
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·”‰…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·”‰…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   6255
         TabIndex        =   18
         Top             =   900
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·”‰…"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "”œœ „Ê”„"
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
         Index           =   4
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "«Œ— „Ê”„ ”œ«œ ›Ì"
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
         Index           =   3
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "·„ Ì”œœ „Ê”„"
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
         Index           =   2
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         Index           =   1
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   960
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   10035
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Key             =   ""
            Object.Tag             =   ""
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   75
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   150
      Top             =   75
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
      Height          =   7665
      Left            =   1665
      TabIndex        =   7
      Top             =   1395
      Width           =   17115
      _cx             =   30189
      _cy             =   13520
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   9840
      Visible         =   0   'False
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdmember1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String, aHeader(), oSearchYear As New Search_empty
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection

Private Sub cmdExel_Click()
ToFileExel grid1, , , prog1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set grdmember1 = Nothing
End Sub
Private Sub CmdUndo_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
myload
End Sub
Private Sub cmdPrint_Click()
Set PrintGrdNew.myForm = Me
PrintGrdNew.doPrint grid1, grdRate(grid1, 15000), 0, "»Ì«‰«  «·«⁄÷«¡", retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, True, 11
PrintGrdNew.Show 1
End Sub
Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT CODE,DESCA FROM TYPE_CODES ORDER BY CODE", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set grid1.DataSource = data10
Fixgrd

grid1.ExplorerBar = flexExSortShow
End Sub
Private Sub myload()
Dim cString As String, cWhere As String, bForm As Boolean
ReDim aHeader(4)

With grid1
cString = "SELECT FILE1_10.CODE AS [«·ﬂÊœ], FILE1_10.DESCA,TYPE_CODES.DESCA,MAX(CASE WHEN FILE6_20H.FORM_NO IS NULL THEN NULL ELSE CONVERT(VARCHAR(10),FILE6_20H.DATE,111) END),MOBIL AS [«·„Ê»Ì·],PHONE AS [«· ·Ì›Ê‰],MAIL AS [«·»—Ìœ «·«·Ìﬂ —Ê‰Ì] " & _
          " FROM FILE1_10 LEFT JOIN TYPE_CODES ON FILE1_10.TYPE = TYPE_CODES.CODE LEFT JOIN FILE6_20H ON dbo.f_last_year_doc(FILE1_10.CODE) = FILE6_20H.DOC_NO"

If xDrop.Value = 0 Then
    cWhere = cWhere & Tr(cWhere) & "FILE1_10.[DROP] = 0"
End If

If xType.MatchedWithList Then
    aHeader(0) = "›∆… «·⁄÷ÊÌ… : " & xType.text
    cWhere = cWhere & turn(cWhere, " AND ") & " file1_10.[TYPE]  = " & addvalue(xType.BoundText)
End If

If IsDate(xDate1.text) Then
    aHeader(1) = BetweenString(xDate1.text, xDate1.text)
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE >= " & DateSq(xDate2.text)
End If
    
If IsDate(xDate2.text) Then
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE <= " & DateSq(xDate2.text)
End If

If IsNumeric(cmdYear(0).Tag) Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.YEAR_CODE >= " & Val(cmdYear(0).Tag)
    aHeader(2) = "”œœ „Ê”„ : " & cmdYear(0).Caption
End If

If IsNumeric(cmdYear(1).Tag) Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.YEAR_CODE = " & Val(cmdYear(1).Tag)
    aHeader(3) = "«Œ— ”œ«œ ›Ì : " & cmdYear(1).Caption
End If

If IsNumeric(cmdYear(2).Tag) Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.YEAR_CODE < " & Val(cmdYear(2).Tag)
    aHeader(4) = "·„ Ì”œœ „Ê”„ : " & cmdYear(1).Caption
End If

If IsDate(xDate1.text) Or IsDate(xDate2.text) Or IsNumeric(cmdYear(0).Tag) Or IsNumeric(cmdYear(1).Tag) Or IsNumeric(cmdYear(2).Tag) Then
    cWhere = cWhere & turn(cWhere, " and ") & "(NOT FORM_NO IS NULL)"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If
cString = cString & " GROUP BY FILE1_10.CODE, FILE1_10.DESCA,TYPE_CODES.DESCA,MOBIL,PHONE,MAIL"

End With
Set data10.Recordset = myCmd(cString, con, , , 500)
Fixgrd
End Sub
Sub Fixgrd()
With grid1
    .RowHeight(0) = 800
    '.WordWrap = True
    '.FrozenCols = 3
    .FormatString = "„|" & "«·ﬂÊœ|" & "«·«”„|" & "›∆… «·⁄÷ÊÌ…|" & " «—ÌŒ «·”œ«œ|" & "—ﬁ„ «·„Ê»Ì·|" & "—ﬁ„ «· ·Ì›Ê‰|" & "«·»—Ìœ «·«·Ìﬂ —Ê‰Ì"
    .ColWidth(0) = 700
    .ColWidth(1) = 1000
    .ColWidth(2) = 2500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1350
    .ColWidth(5) = 2500
    .ColWidth(6) = 2500
    .ColWidth(7) = 2500
    
    .ExplorerBar = flexExSortAndMove
    .Cell(flexcpAlignment, 0, 0, .rows - 1, .Cols - 1) = 4
     StatusBar1.Panels(1).text = IIf(grid1.rows > 1, "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.rows - 1, "")
    .ExplorerBar = flexExSort
    For i = 1 To grid1.rows - 1
        .TextMatrix(i, 0) = i
    Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 6)
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(0).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·”‰…", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub

