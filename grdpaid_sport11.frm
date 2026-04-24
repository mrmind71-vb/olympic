VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdpaid_sport1 
   Caption         =   "ÅÌãÇáí ÓÏÇÏ äÔÇØ ÕíÝí ÎáÇá ÝÊÑÉ"
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   4830
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3645
         Picture         =   "grdpaid_sport11.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   2415
         Picture         =   "grdpaid_sport11.frx":3059
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdpaid_sport11.frx":5483
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "grdpaid_sport11.frx":78EF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "ÚÑÖ"
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
      Begin MSDataListLib.DataCombo xuser 
         Height          =   330
         Left            =   2520
         TabIndex        =   12
         Top             =   990
         Width           =   2445
         _ExtentX        =   4313
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
      Begin VB.Label Label12 
         Caption         =   "ÇáãÓÊÎÏã"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1035
         Width           =   1095
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
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
      Height          =   7620
      Left            =   135
      TabIndex        =   5
      Top             =   1485
      Width           =   16620
      _cx             =   29316
      _cy             =   13441
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
End
Attribute VB_Name = "grdpaid_sport1"
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
Dim cHead1 As String, cHead2 As String, cHead3 As String
cHead1 = "ÅÌãÇáì ÓÏÇÏ ÎáÇá ÝÊÑÉ"
cHead2 = retHeader(aHeader, 0, 1)
cHead3 = retHeader(aHeader, 1, 1)
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)
PrintGrdNew.doprint grid1, 1, -2, cHead1, cHead2, cHead3, , False, True, 9, , aRow
PrintGrdNew.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
MyLoad
End Sub
Private Sub Form_Load()
Me.Top = 1000
Me.Left = 1000
openCon con


Set data1.Recordset = myRecordSet("select * from users", con)
Set xuser.RowSource = data1
xuser.ListField = "Desca"
xuser.BoundColumn = "Code"


Set grid1.DataSource = data10
data10.ConnectionString = strCon
Fixgrd
LoadText Me
End Sub
Private Sub MyLoad()
Dim cString As String, nPrevious As Double
ReDim aHeader(2)
With grid1
If Not MYVALID Then Exit Sub
cString = "SELECT '', FILE6_30H.DOC_NO,FILE6_30H.[NAME], CONVERT(VARCHAR(10),FILE6_30H.DATE,111),SUM(FILE6_30.TOTAL),'' AS ITEMS_DESCA " & _
          "  FROM FILE6_30H INNER JOIN FILE6_30 ON FILE6_30H.DOC_NO = FILE6_30.DOC_NO"

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "FILE6_30H.DATE >= " & DateSq(xDate1.Text)
    aHeader(1) = BetweenString(xDate1.Text, xDate2.Text)
End If
          
If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "FILE6_30H.DATE <= " & DateSq(xDate2.Text)
    aHeader(1) = BetweenString(xDate1.Text, xDate2.Text)
End If

If xuser.MatchedWithList Then
    cString = cString & turn(cString) & "FILE6_30H.USERCODE = " & xuser.BoundText
    aHeader(2) = "ÇáãæÙÝ : " & xuser.Text
End If

cString = cString & " group by FILE6_30H.DOC_NO,FILE6_30H.[NAME],FILE6_30H.DATE Order by FILE6_30H.DOC_NO"
data10.RecordSource = cString
data10.Refresh
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
    .TextMatrix(0, 2) = "ÇÓã ÇáÚÖæ"
    .TextMatrix(0, 3) = "ÊÇÑíÎ ÇáãÓÊäÏ"
    .TextMatrix(0, 4) = "ÞíãÉ ÇáãÓÊäÏ"
    .TextMatrix(0, 5) = "ÇáÈíÇä"
        
    .ColWidth(0) = 900
    .ColWidth(1) = 1000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1350
    .ColWidth(4) = 1000
    .ColWidth(5) = 10000
    
    .ColHidden(0) = True

'    .ColDataType(2) = flexDTDouble
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble

    
    For I = 0 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 1 To grid1.Rows - 1
        .TextMatrix(I, 0) = I
    Next
               
               
    If .Rows > 1 Then
        .SubtotalPosition = flexSTAbove
        .Subtotal flexSTSum, -1, 4, "#0.00", vbRed, vbYellow, True, "  "
        For I = 0 To 3
            .TextMatrix(1, I) = "ÇáÅÌãÇáì"
        Next
        .MergeCells = flexMergeFree
        .MergeRow(1) = True
    End If
    Dim aret As Variant, cString As String, cField As String
    For I = 2 To .Rows - 1
        cString = "SELECT  FILE1_30G.DESCA,SUM(QUANT) AS countOf FROM   FILE6_30 INNER JOIN FILE1_30 ON FILE6_30.CODE = FILE1_30.CODE INNER JOIN FILE1_30G ON FILE1_30.[GROUP] = FILE1_30G.CODE "
        cString = cString & turn(cString) & "FILE6_30.DOC_NO = " & MyParn(grid1.TextMatrix(I, 1))
        cString = cString & "GROUP BY FILE1_30G.DESCA"
        aret = GetRows(cString, con)
        cField = ""
        If Not IsEmpty(aret) Then
            For i2 = 0 To UBound(aret)
                cField = cField & turn(cField, "-") & retFlag(aret(i2), "desca") & " (" & retFlag(aret(i2), "countOf") & ")"
            Next
        End If
        grid1.TextMatrix(I, 5) = cField
    Next
    
    SBar1.Panels(1).Text = IIf(grid1.Rows > 2, "ÚÏÏ ÇáÓÌáÇÊ : " & grid1.Rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdbankfrm1 = Nothing
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
myLostFocus XCODE
xCodeDesca.Caption = ""
If Not ValidInt(XCODE.Text) Then Exit Sub
Dim aret As Variant
aret = GetFields("select DESCA from file1_10 where code = " & XCODE.Text)
If Not IsEmpty(aret) Then xCodeDesca.Caption = retFlag(aret, "DESCA") & ""
End Sub

Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
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
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xbox_GotFocus()
myGotFocus xbox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
End Sub
Private Sub xCode_GotFocus()
myGotFocus XCODE
End Sub

Sub myProc()
XCODE.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
Unload oSearchMember
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub xShare_Click(Area As Integer)

End Sub
