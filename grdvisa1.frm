VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdVisafrm1 
   Caption         =   "„ «»⁄…  ›’Ì·Ì ›Ì“«"
   ClientHeight    =   9705
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   11940
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
   ScaleHeight     =   9705
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   540
      Width           =   4830
      Begin VB.CommandButton cmdprint2 
         Height          =   510
         Left            =   1230
         Picture         =   "grdvisa1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   510
         Left            =   2430
         Picture         =   "grdvisa1.frx":2462
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   3600
         Picture         =   "grdvisa1.frx":488C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "grdvisa1.frx":78E5
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   6855
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬂ«‘Ì— :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   1215
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   495
      Top             =   -135
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
      Left            =   1710
      Top             =   -270
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
      Left            =   990
      Top             =   -270
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   900
      Top             =   -135
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
      Height          =   8565
      Left            =   90
      TabIndex        =   12
      Top             =   1260
      Width           =   11715
      _cx             =   20664
      _cy             =   15108
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
      Cols            =   6
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
   Begin VB.Label xTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   135
      Width           =   4785
   End
End
Attribute VB_Name = "grdVisafrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim aHeader(1)
Dim oSearchClient As New Search3
Dim cStr1 As String, cStr2 As String
Private Sub cmdPrint_Click()
    Dim cHead1 As String, cHead2 As String, cHead3 As String
    cHead1 = "„ «»⁄… ›Ì“« Œ·«· › —…"
    cHead2 = retHeader(aHeader, 0, 1)
    cHead3 = retHeader(aHeader, 1, 1)
    PrintGrd.doprint Me.grid1, 1, -2, cHead1, cHead2, cHead3, False, False, 8, , Array(1), Array(1)
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
MyLoad
End Sub

Private Sub CmdPrint2_Click()
doprint
End Sub

Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From FILE0_50 ORDER BY DESCA"
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Fixgrd
grid1.Rows = 1
End Sub
Private Sub MyLoad()
Dim cwhere As String
cwhere = ""
For i = 0 To UBound(aHeader)
    aHeader(i) = ""
Next
With grid1

    cString = "SELECT FILE6_20H.DOC_NO,CONVERT(VARCHAR(10),[DATE],111),CONVERT(VARCHAR(5),FILE6_20H.TIME,108),FILE0_50.DESCA,SUM(FILE6_20.TOTAL) - FILE6_20H.DISCOUNT,FILE6_20H.VISA " & _
              " FROM  (FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO) INNER JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE"
    cString = cString & turn(cString) & "FILE6_20H.PRINTED = 1"
    cString = cString & turn(cString) & "FILE6_20H.VISA <> 0"
    If IsDate(xdate1.Text) Then
        cString = cString & turn(cString) & " FILE6_20H.DATE  >= " & DateSq(xdate1.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "„‰ : " & Format(xdate1.Text, "yyyy/mm/dd")
    End If
       
    If IsDate(xDate2.Text) Then
        cString = cString & turn(cString) & " FILE6_20H.DATE  <= " & DateSq(xDate2.Text)
        aHeader(0) = aHeader(0) & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.Text, "yyyy/mm/dd")
    End If
       
    If xBox.BoundText <> "" Then
         cString = cString & turn(cString) & " FILE6_20H.[BOX]  = " & MyParn(xBox.BoundText)
         aHeader(1) = "«·ﬂ«‘Ì— :" & xBox.Text
    End If
    cString = cString & " GROUP BY  FILE6_20H.DOC_NO,FILE6_20H.DATE,FILE6_20H.TIME,FILE6_20H.CASH,FILE6_20H.VISA,FILE6_20H.DISCOUNT,FILE0_50.DESCA"
    cString = cString & " ORDER BY FILE6_20H.DATE,FILE6_20H.DOC_NO"
    DATA10.RecordSource = cString
    DATA10.Refresh
End With
Handlecontrols
Fixgrd
End Sub
Sub Fixgrd()
Dim nTotal As Double, nSales As Double
    With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "—ﬁ„ «·›« Ê—…"
    .TextMatrix(0, 1) = " «—ÌŒ «·›« Ê—…"
    .TextMatrix(0, 2) = "Êﬁ  «·›« Ê—…"
    .TextMatrix(0, 3) = "«·ﬂ«‘Ì—"
    .TextMatrix(0, 4) = "≈Ã„«·Ì «·›« Ê—…"
    .TextMatrix(0, 5) = "«·›Ì“«"

        
    
    .ColWidth(1) = 1600
    .ColWidth(2) = 1900
    .ColWidth(3) = 1300
    .ColWidth(4) = 1300
    
   
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble

    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 4, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0.00", vbRed, vbYellow, True, "  "
    If grid1.Rows > 1 Then
        .MergeCells = flexMergeFree
        For i = 0 To 3
            .TextMatrix(1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(1) = True
    End If
    xtotal.Caption = IIf(grid1.Rows > 2, grid1.Rows - 2, "")
    xtotal.Caption = turn(xtotal.Caption, "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : ") & xtotal.Caption
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDesca.Text, 1
End If
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xItem.Text, 0
End If
End Sub
Private Sub grid1_dblClick()
Dim aHeader As Variant
Dim sHead1 As String, sHead2 As String, sHead3 As String
With grid1
If grid1.Rows < 3 And grid1.Row < 2 Then Exit Sub
    showfrm4.sHead1 = "»Ì«‰«  «·›« Ê—… —ﬁ„ : " & .TextMatrix(grid1.Row, 0)
    showfrm4.sDoc_No = .TextMatrix(grid1.Row, 0)
    showfrm4.Show 1
End With
End Sub
Sub myProc()
ActiveControl.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
Unload oSearchClient
End Sub
Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text)
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xCode.Text)) & ""
End Sub
Private Sub xSection_Validate(Cancel As Boolean)
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xbox_Validate(Cancel As Boolean)
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub
Private Sub xGroup2_Validate(Cancel As Boolean)
If Not xGroup2.MatchedWithList Then xGroup2.BoundText = ""
End Sub
Private Sub Xcode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
xCode.BackColor = &HC0FFFF
End Sub
Private Sub xdate1_GotFocus()
xdate1.SelStart = 0
xdate1.SelLength = Len(xdate1.Text)
xdate1.BackColor = &HC0FFFF
End Sub
Private Sub xDate2_GotFocus()
xDate2.SelStart = 0
xDate2.SelLength = Len(xDate2.Text)
xDate2.BackColor = &HC0FFFF
End Sub
Private Sub xgroup_GotFocus()
xgroup.BackColor = &HC0FFFF
End Sub
Private Sub xSection_GotFocus()
xSection.BackColor = &HC0FFFF
End Sub
Private Sub xBox_GotFocus()
xBox.BackColor = &HC0FFFF
End Sub
Private Sub xGroup2_GotFocus()
xGroup2.BackColor = &HC0FFFF
End Sub
Private Sub xdate1_LostFocus()
xdate1.BackColor = &H80000005
End Sub
Private Sub xDate2_LostFocus()
xDate2.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xgroup.BackColor = &H80000005
End Sub
Private Sub xSection_LostFocus()
xSection.BackColor = &H80000005
End Sub
Private Sub xBox_LostFocus()
xBox.BackColor = &H80000005
End Sub
Private Sub xgroup2_LostFocus()
xGroup2.BackColor = &H80000005
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
End Sub
Private Function doprint()
'On Error GoTo myerror
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
     
If IsDate(xdate1.Text) Then aHeader(0) = BetweenString(xdate1.Text, xDate2.Text)
If IsDate(xDate2.Text) Then aHeader(0) = BetweenString(xdate1.Text, xDate2.Text)
If xBox.MatchedWithList Then aHeader(1) = turn(xBox.Text, "«·ﬂ«‘Ì— : ") & xBox.Text

With grid1
For i = 2 To grid1.Rows - 1
    
    temptable.AddNew
    temptable!str1 = TurnValue(ArbString(grid1.TextMatrix(i, 1)))
    temptable!str2 = TurnValue(grid1.TextMatrix(i, 2))
    temptable!val1 = Val(grid1.TextMatrix(i, 5))
    temptable!str12 = Format(grid1.TextMatrix(i, 1), "yyyy-mm-dd")
    temptable!Str11 = arbDay(grid1.TextMatrix(i, 1)) & turn(grid1.TextMatrix(i, 1), " ") & Format(grid1.TextMatrix(i, 1), "yyyy/mm/dd")
    temptable!str21 = TurnValue(aHeader(0))
    If Trim(aHeader(1)) <> "" Then temptable!str21 = temptable!str21 & turn(temptable!str21, vbCrLf) & aHeader(1)
    temptable.Update
Next
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    temptable.Requery
    main.REPORT1.Reset
    'Report1.PrinterCopies = nCopies
    'Report1.Destination = crptToPrinter
    main.REPORT1.ReportFileName = App.Path & "\Reports\visa1.rpt"
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If
doprint = True
closeCon:
temptable.Close
Set temptable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function
Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

