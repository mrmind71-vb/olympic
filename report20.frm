VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm20 
   Caption         =   "Œÿ«»«  «‰–«—"
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
      Height          =   1320
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   90
      Width           =   6990
      Begin VB.TextBox xDate_End 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "D"
         Top             =   900
         Width           =   1410
      End
      Begin VB.TextBox xPeriod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   180
         Width           =   2670
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox xMeeting 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox xDate_Send 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «·«‰ Â«¡"
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
         Index           =   5
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·„œ…"
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
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «·Ã·”…"
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·Ã·”…"
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «· Õ—Ì—"
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
         Index           =   0
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   675
      Width           =   5325
      Begin VB.CommandButton cmdPrintLetter 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   3195
         Picture         =   "report20.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4245
         Picture         =   "report20.frx":24FE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2145
         Picture         =   "report20.frx":49F0
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "report20.frx":6E1A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1095
         Picture         =   "report20.frx":9286
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
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
      Height          =   960
      Left            =   12420
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   450
      Width           =   4290
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Õ«›Ÿ ⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   585
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ ⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1365
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   1620
         TabIndex        =   0
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "«Œ— ”œ«œ ›Ì"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2340
      Top             =   450
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
      Height          =   7395
      Left            =   90
      TabIndex        =   9
      Top             =   1440
      Width           =   16620
      _cx             =   29316
      _cy             =   13044
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
      Cols            =   9
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
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "reportfrm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Dim aHeader()

Private Sub cmdExel_Click()
Dim sHeader As String, nMargin As Integer

sHeader = Me.Caption
nMargin = 40
If retHeader(aHeader, 1, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 1, 3)
    nMargin = nMargin + 15
End If



Dim aSplit As Variant
aSplit = AddFlag(aSplit, "title_col", "A:B")
aSplit = AddFlag(aSplit, "title_row", "1:1")
aSplit = AddFlag(aSplit, "center_header", sHeader)
ToFileExel grid1, , , , , , , , aSplit, , , , nMargin
End Sub
Private Sub cmdPrint_Click()
PrintGrdNew.doprint grid1, 0.85, 0, Me.Caption, retHeader(aHeader, 1, 2), retHeader(aHeader, 3, 2), , False, False, 10
PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
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
Private Sub cmdPrintLetter_Click()
doprint
End Sub

Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub
Private Sub Form_Load()
MFocus Me
Me.Top = 1000
Me.Left = 1000
openCon con

Set grid1.DataSource = data10
Fixgrd
LoadText Me
If Not IsDate(xDate_Send.text) Then xDate_Send.text = myFormat_p(Date)
If ValidNum(cmdYear(1).Tag) Then
    cmdYear(1).Caption = GetField("select desca from years_codes where code = " & cmdYear(1).Tag, con) & ""
End If
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(4)
With grid1

cString = "SELECT FILE1_10.CODE,FILE1_10.DESCA,FILE1_10.TITLE,FILE1_10.ADDRESS,FILE1_10.GENDER,dbo.f_last_year_date(file1_10.code) as date_last,dbo.f_last_year_doc(file1_10.code) as doc_no_last,dbo.f_save(FILE1_10.CODE) AS IS_SAVE" & _
          " FROM FILE1_10"
cWhere = "FILE1_10.[DROP] = 0"

cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(file1_10.code) = " & cmdYear(1).Tag

If Check1.Value = 1 Then
    cWhere = cWhere & turn(cWhere, " AND ") & "ADDRESS IS NULL"
    aHeader(1) = "»œÊ‰ ⁄‰Ê«‰"
End If

If Check2.Value = 1 Then
    cWhere = cWhere & turn(cWhere, " AND ") & "dbo.f_save(FILE1_10.CODE) = 1"
    aHeader(2) = "Õ«›ŸÌ ⁄÷ÊÌ…"
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If cmdYear(1).Tag <> "" Then
    aHeader(3) = "··«⁄÷«¡ «·–Ì‰ «Œ— ”œ«œ ·Â„ ›Ï „Ê”„ : " & cmdYear(1).Caption
End If

If IsDate(xDate.text) Then
    aHeader(4) = "„‰  «—ÌŒ : " & myFormat_p(xDate.text)
End If

aHeader(0) = getYears(cmdYear(1).Tag, xDate.text)

cString = cString & " ORDER BY FILE1_10.CODE"
Set data10.Recordset = myCmd(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim I As Long
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 2) = "«”„ «·⁄÷Ê"
    .TextMatrix(0, 3) = "«··ﬁ»"
    .TextMatrix(0, 4) = "«·⁄‰Ê«‰"
    .TextMatrix(0, 5) = "«·‰Ê⁄"
    .TextMatrix(0, 6) = " «—ÌŒ ”œ«œ"
    
    .TextMatrix(0, 7) = "„” ‰œ ”œ«œ"
    .TextMatrix(0, 8) = "Õ«›Ÿ ⁄÷ÊÌ…"
        
    .ColWidth(0) = 800
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1800
    .ColWidth(4) = 5000
    .ColWidth(5) = 1300
    .ColWidth(6) = 1350
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColHidden(3) = True
    .ColHidden(5) = True
    .ColDataType(6) = flexDTDate
    .ColDataType(8) = flexDTBoolean
    
    For I = 1 To grid1.rows - 1
        .TextMatrix(I, 0) = I
        .TextMatrix(I, 6) = myFormat_p(.TextMatrix(I, 6))
    Next
    
    For I = 0 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    SBar1.Panels(1).text = IIf(grid1.rows > 2, "⁄œœ «·”Ã·«  : " & grid1.rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me, , Array(xDate.Name, xDate_Send.Name, xMeeting.Name, cmdYear(1).Name, xDate_End.Name, xPeriod.Name)
closeCon con
Set grdpaid3 = Nothing
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub
Private Sub grid1_DblClick()
With grid1
End With
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidInt(xCode.text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode.text)
If Not IsEmpty(aRet) Then xCodeDesca.Caption = retFlag(aRet, "DESCA") & ""
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
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
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
If ActiveControl.Name = cmdYear(1).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Function MYVALID() As Boolean
If cmdYear(1).Tag = "" Then
    MsgBox "·« ÌÊÃœ „Ê”„"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox IIf(Trim(xDate.text) = "", " «—ÌŒ €Ì— „”Ã·", " «—ÌŒ €Ì— ’ÕÌÕ")
    Exit Function
End If

If Trim(xMeeting.text) = "" Then
    MsgBox "—ﬁ„ «·Ã·”… €Ì— „”Ã·"
    Exit Function
End If


If Not IsDate(xDate_Send.text) Then
    MsgBox IIf(Trim(xDate_Send.text) = "", " «—ÌŒ «·Ã·”… €Ì— „”Ã·", " «—ÌŒ «·Ã·”… €Ì— ’ÕÌÕ")
    Exit Function
End If

MYVALID = True
End Function
Private Function getYears(pCode As String, pDate As String) As String
Dim loctable As New ADODB.Recordset
cString = "select desca_r from years_codes where code <=  dbo.f_ret_year(" & DateSq(pDate) & ") and code > " & cmdYear(1).Tag & "  group by [desca_r],code order by code"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    getYears = getYears & turn(getYears, " ° ") & loctable!Desca_r
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function
Private Function doprint()
Dim temptable As New ADODB.Recordset, cOr As String

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim sString As String
With grid1
For I = 1 To grid1.rows - 1
    If Trim(.TextMatrix(I, 4)) <> "" Then
        temptable.AddNew
        temptable!val1 = TurnValue(.TextMatrix(I, 1))
        temptable!str1 = TurnValue(ArbString(.TextMatrix(I, 1)))
        temptable!str2 = TurnValue(.TextMatrix(I, 2))
        temptable!str5 = " «—ÌŒ «· Õ—Ì— : " & myFormat_p(xDate_Send.text)
        temptable!str16 = TurnValue(ArbString(.TextMatrix(I, 4)))
        sString = "‰ÕÌÿ ”Ì«œ ﬂ„ ⁄·„« »√‰ „Ã·” «œ«—… «·‰«œÌ »Ã·”… —ﬁ„ " & xMeeting.text & " » «—ÌŒ " & myFormat_p(xDate.text)
        sString = sString & " " & "ﬁœ  ﬁ—— „‰Õﬂ„ „Â·… " & xPeriod.text
        sString = sString & " " & "··„»«œ—… »”œ«œ «·«‘ —«ﬂ «·„ √Œ— ⁄·Ìﬂ„ ⁄‰ «·„Ê«”„ : " & vbCrLf
        sString = sString & " " & " ( " & aHeader(0) & " ) "
        If IsDate(xDate_End.text) Then
            sString = sString & vbCrLf & " ÕÌÀ  ‰ ÂÌ «·„Â·… ›Ï " & myFormat_p(xDate_End.text)
        End If
        'temptable!memo1 = Trim(Chr(254) & Replace(sString, " ", Chr(254) & " " & Chr(254)))
        temptable!memo1 = ArbString(sString)
        temptable.Update
    End If
Next
End With

temptable.Requery
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    Report1.ReportFileName = sPath_App & "\REPORTS\warrent.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
End Function
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_Send_GotFocus()
myGotFocus xDate_Send
End Sub
Private Sub xDate_Send_LostFocus()
myLostFocus xDate_Send
myValidDate xDate_Send
End Sub
Private Sub xDate_end_GotFocus()
myGotFocus xDate_End
End Sub
Private Sub xDate_end_LostFocus()
myLostFocus xDate_End
myValidDate xDate_End
End Sub
Private Sub xPeriod_GotFocus()
myGotFocus xPeriod
End Sub
Private Sub xPeriod_LostFocus()
myLostFocus xPeriod
End Sub
Private Sub xMeeting_GotFocus()
myGotFocus xMeeting
End Sub
Private Sub xMeeting_LostFocus()
myLostFocus xMeeting
End Sub
