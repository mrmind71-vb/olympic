VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grdTrust1 
   Caption         =   "≈Ã„«·Ì „Êﬁ› œ«∆‰Ì‰"
   ClientHeight    =   10140
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10140
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   405
      Width           =   5595
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1140
         Picture         =   "grdtrust12.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2235
         Picture         =   "grdtrust12.frx":2424
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4425
         Picture         =   "grdtrust12.frx":4C0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdtrust12.frx":7101
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3330
         Picture         =   "grdtrust12.frx":956D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   9855
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   5280
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   3795
         _ExtentX        =   6694
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
      Begin VB.Label xBal_Done 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "—’Ìœ  ”ÊÌ« "
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "«·”«∆ﬁ"
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
         Index           =   3
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   600
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   9810
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data11 
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
   Begin MSAdodcLib.Adodc DATA1 
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
      Height          =   8565
      Left            =   270
      TabIndex        =   0
      Top             =   1170
      Width           =   14865
      _cx             =   26220
      _cy             =   15108
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
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
End
Attribute VB_Name = "grdTrust1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection

Private Sub CmdClear_Click()
DefineText Me
End Sub

Private Sub cmdExit_Click()
Unload Me
Set grdDebit1 = Nothing
End Sub
Private Sub CmdUndo_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
If Not xDriver.MatchedWithList Then
    MsgBox "ﬂÊœ «·”Ê«ﬁ €Ì— „”Ã·"
    Exit Sub
End If
myload
End Sub
Private Sub CmdPrint_Click()
Dim cHeader1 As String
Dim aHeader As Variant
cHeader1 = "»Ì«‰  ”ÊÌ«  «·”«∆ﬁÌ‰ Œ·«· › —…"
If xDriver.MatchedWithList Then aHeader = AddFlag(aHeader, "«·”«∆ﬁ : " & xDriver.Text)
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 2)
PrintGrdNew.doprint grid1, 0.9, -2, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 9, , aRow
PrintGrdNew.Show 1
End Sub
Private Sub Form_Load()
openCon con
Set data1.Recordset = myRecordSet("Select driver.Code,driver.DescA From driver inner join trust_h on driver.code = trust_h.box where driver.driver = 1 order by driver.Desca", con)
Set xDriver.RowSource = data1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

Set grid1.DataSource = data11
data11.ConnectionString = strCon
Fixgrd
grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String, aCharge As Variant
cString = "Select -1 * sum(TRUST_BAL1.TOTAL) FROM TRUST_BAL1 WHERE BOX = " & MyParn(xDriver.BoundText)
xBal_Done.Caption = Myvalue(GetField(cString))
cString = "SELECT DISTINCT FILE8_51.CODE, FILE8_51.DESCA" & _
          " FROM TRAVEL_C INNER JOIN FILE8_51 ON TRAVEL_C.Charge = FILE8_51.CODE" & _
          " INNER JOIN TRAVEL_NOT_DONE ON TRAVEL_C.DOC_NO = TRAVEL_NOT_DONE.TRAVEL" & _
          "  WHERE TRAVEL_C.BOX = " & MyParn(xDriver.BoundText)
aCharge = GetRows(cString)

If Not IsEmpty(aCharge) Then
    For I = 0 To UBound(aCharge)
        cField = cField & turn(cField, ",") & _
                myiif("dbo.TRAVEL_C.Charge = " & MyParn(retFlag(aCharge(I), "CODE")), "TRAVEL_C.VALUE") & " AS " & "[" & retFlag(aCharge(I), "DESCA") & "]"
    Next
End If

With grid1
cString = "SELECT dbo.TRUST_H.DOC_NO,TRUST_H.DATE, dbo.TRUST_TOTAL.TRUST_TOTAL" & turn(cField, ",") & cField & _
          " FROM  dbo.TRUST_H INNER JOIN " & _
          " dbo.TRUST_DOC ON dbo.TRUST_H.DOC_NO = dbo.TRUST_DOC.DOC_NO LEFT JOIN" & _
          " dbo.TRAVEL_C ON (dbo.TRUST_DOC.TRAVEL = dbo.TRAVEL_C.DOC_NO AND TRAVEL_C.BOX = " & MyParn(xDriver.BoundText) & ")" & _
          " LEFT OUTER JOIN " & _
          " dbo.TRUST_TOTAL ON (dbo.TRUST_H.DOC_NO = dbo.TRUST_TOTAL.DOC_NO AND TRUST_TOTAL.BOX = " & MyParn(xDriver.BoundText) & ")"
cString = cString & turn(cString) & "TRUST_H.DONE = 0"
cString = cString & turn(cString) & "TRUST_H.BOX = " & MyParn(xDriver.BoundText)
cString = cString & " GROUP BY dbo.TRUST_H.DOC_NO,TRUST_H.DATE, dbo.TRUST_TOTAL.TRUST_TOTAL"
data11.RecordSource = cString
data11.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
Dim I As Long
With grid1
.RowHeight(0) = 1000
.WordWrap = True
.FrozenCols = 2
.TextMatrix(0, 0) = "—ﬁ„ «· ”ÊÌ…"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·⁄Âœ…"
If .Cols > 3 Then
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = "«·—’Ìœ"
End If

For I = 2 To .Cols - 2
    .ColWidth(I) = 1000
Next
.ColWidth(.Cols - 1) = 1200
.MergeCells = flexMergeFree
For I = 2 To .Cols - 1
    .ColDataType(I) = flexDTDouble
Next

If .Cols > 3 Then
    Dim nTotal_Charge As Double, nRow As Long
    For nRow = 1 To .Rows - 1
        nTotal_Charge = 0
        For I = 3 To .Cols - 2
            nTotal_Charge = nTotal_Charge + Val(.TextMatrix(nRow, I))
        Next
        .TextMatrix(nRow, .Cols - 1) = nTotal_Charge - Val(.TextMatrix(nRow, 2))
    Next
End If

.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
.SubtotalPosition = flexSTAbove
For I = 2 To .Cols - 1
    .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
Next
StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
If .Rows > 1 Then
    .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
    .TextMatrix(1, 1) = "«·≈Ã„«·Ì"
    .MergeRow(1) = True
End If
.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then CardLookup
End Sub
