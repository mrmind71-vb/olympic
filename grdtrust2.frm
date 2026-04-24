VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grdTrust2 
   Caption         =   " ”ÊÌ… ”«∆ﬁÌ‰"
   ClientHeight    =   10140
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   19950
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
   ScaleWidth      =   19950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "«Œ Ì«—« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   315
      Width           =   6900
      Begin VB.OptionButton Option3 
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
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   315
         Value           =   -1  'True
         Width           =   1950
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " „ ⁄„·  ”ÊÌ…"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   315
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "·„ Ì „ ⁄„·  ”ÊÌ…"
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   315
         Width           =   1950
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   315
      Width           =   5370
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1095
         Picture         =   "grdtrust2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2145
         Picture         =   "grdtrust2.frx":2424
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4230
         Picture         =   "grdtrust2.frx":4C0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   45
         Picture         =   "grdtrust2.frx":7101
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3195
         Picture         =   "grdtrust2.frx":956D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   14490
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   5280
      Begin VB.TextBox xTrust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2025
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   135
         TabIndex        =   0
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
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «· ”ÊÌ…"
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
         TabIndex        =   12
         Top             =   585
         Width           =   1050
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
         TabIndex        =   11
         Top             =   270
         Width           =   600
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   9810
      Width           =   19950
      _ExtentX        =   35190
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
      Height          =   465
      Left            =   1080
      Top             =   -315
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
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
      Left            =   855
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
      Left            =   90
      TabIndex        =   2
      Top             =   1080
      Width           =   19770
      _cx             =   34872
      _cy             =   14473
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
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9180
      Width           =   19725
      Begin ComctlLib.ProgressBar prog1 
         Height          =   330
         Left            =   45
         TabIndex        =   18
         Top             =   180
         Visible         =   0   'False
         Width           =   18240
         _ExtentX        =   32173
         _ExtentY        =   582
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "grdTrust2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection, oSearchDriver As New Search3
Dim oSearch As New Search3
Private Sub CmdClear_Click()
DefineText Me
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1
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
Private Sub cmdPrint_Click()
Dim cHeader1 As String
Dim aHeader As Variant
cHeader1 = "»Ì«‰  ”ÊÌ«  «·”«∆ﬁÌ‰ Œ·«· › —…"
If xDriver.MatchedWithList Then aHeader = AddFlag(aHeader, "«·”«∆ﬁ : " & xDriver.Text)
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 2)
Set PrintGrdNew.myForm = Me
PrintGrdNew.doprint grid1, 0.8, -1, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, True, 8, , aRow
PrintGrdNew.Show 1
End Sub
Private Sub Form_Load()
openCon con
Set DATA1.Recordset = myRecordSet("Select driver.Code,driver.DescA From driver where driver.driver = 1 order by driver.Desca", con)
Set xDriver.RowSource = DATA1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

Set grid1.DataSource = data11
data11.ConnectionString = strCon
Fixgrd
grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String, aCharge As Variant
'cString = "Select -1 * sum(TRUST_BAL1.TOTAL) FROM TRUST_BAL1 WHERE BOX = " & MyParn(xDriver.BoundText)
'xBal_Done.Caption = Myvalue(GetField(cString))

cString = "SELECT DISTINCT FILE8_51.CODE, FILE8_51.DESCA" & _
          " FROM TRAVEL_C INNER JOIN FILE8_51 ON TRAVEL_C.Charge = FILE8_51.CODE" & _
          " LEFT JOIN V_TRUST_DOC ON (TRAVEL_C.DOC_NO = V_TRUST_DOC.TRAVEL " & _
          " AND V_TRUST_DOC.BOX = " & MyParn(xDriver.BoundText) & ")" & _
          " WHERE TRAVEL_C.BOX = " & MyParn(xDriver.BoundText)
          
If Trim(xTrust.Text) <> "" Then
    cString = cString & turn(cString) & " V_TRUST_DOC.DOC_NO = " & MyParn(xTrust.Text)
End If
If Option1.Value Then
    cString = cString & turn(cString) & "V_TRUST_DOC.DOC_NO IS NULL"
ElseIf Option2.Value Then
    cString = cString & turn(cString) & "(NOT (V_TRUST_DOC.DOC_NO IS NULL))"
End If
aCharge = GetRows(cString)
If Not IsEmpty(aCharge) Then
    For i = 0 To UBound(aCharge)
        cField = cField & turn(cField, ",") & _
                myiif("dbo.TRAVEL_C.Charge = " & MyParn(retFlag(aCharge(i), "CODE")), "TRAVEL_C.VALUE") & " AS " & "[" & retFlag(aCharge(i), "DESCA") & "]"
    Next
End If

With grid1
cString = "SELECT V_TRUST_DOC.DOC_NO,CONVERT(VARCHAR(10),dbo.TRAVEL_H.DATE,111),dbo.TRAVEL_H.DOC_NO,CARS.BOARD,PLACE_CODES.DESCA,PLACE_CODES_1.DESCA,CARGO_CODES.DESCA,TRAVEL_BAL.TRUST" & turn(cField, ",") & cField & _
          " FROM TRAVEL_H LEFT JOIN " & _
          " TRAVEL_C ON (dbo.TRAVEL_H.DOC_NO = dbo.TRAVEL_C.DOC_NO AND TRAVEL_C.BOX = " & MyParn(xDriver.BoundText) & ")" & _
          " LEFT OUTER JOIN " & _
          " TRAVEL_BAL ON (dbo.TRAVEL_BAL.DOC_NO = TRAVEL_H.DOC_NO AND TRAVEL_BAL.BOX = " & MyParn(xDriver.BoundText) & ")" & _
          " LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE" & _
          " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE" & _
          " LEFT JOIN PLACE_CODES AS PLACE_CODES_1 ON TRAVEL_H.PLACE2 = PLACE_CODES_1.CODE" & _
          " LEFT JOIN CARGO_CODES ON TRAVEL_H.CARGO = CARGO_CODES.CODE" & _
          " LEFT JOIN V_TRUST_DOC ON (TRAVEL_H.DOC_NO = V_TRUST_DOC.TRAVEL " & _
          " AND V_TRUST_DOC.BOX = " & MyParn(xDriver.BoundText) & ")"
If Trim(xTrust.Text) <> "" Then
    cString = cString & turn(cString) & " V_TRUST_DOC.DOC_NO = " & MyParn(xTrust.Text)
End If
If Option1.Value Then
    cString = cString & turn(cString) & "V_TRUST_DOC.DOC_NO IS NULL"
ElseIf Option2.Value Then
    cString = cString & turn(cString) & "(NOT (V_TRUST_DOC.DOC_NO IS NULL))"
End If

cString = cString & turn(cString) & "( NOT (TRAVEL_C.DOC_NO IS NULL AND TRAVEL_BAL.TRUST = 0 ) )"
cString = cString & " GROUP BY V_TRUST_DOC.DOC_NO,dbo.TRAVEL_H.DOC_NO,TRAVEL_H.DATE,CARS.BOARD,PLACE_CODES.DESCA,PLACE_CODES_1.DESCA,CARGO_CODES.DESCA,TRAVEL_BAL.TRUST"
cString = cString & " ORDER BY V_TRUST_DOC.DOC_NO,dbo.TRAVEL_H.DATE,dbo.TRAVEL_H.DOC_NO"
data11.RecordSource = cString
data11.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
Dim i As Long
With grid1
.RowHeight(0) = 1000
.WordWrap = True
.FrozenCols = 3
.TextMatrix(0, 0) = "«· ”ÊÌ…"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "—ﬁ„ «·—Õ·…"
.TextMatrix(0, 3) = "—ﬁ„ «·”Ì«—…"
.TextMatrix(0, 4) = "„‰"
.TextMatrix(0, 5) = "≈·Ï"
.TextMatrix(0, 6) = "«·Õ„Ê·…"
.TextMatrix(0, 7) = "«·⁄Âœ…"
If .Cols > 8 Then
    .Cols = .Cols + 2
    .TextMatrix(0, .Cols - 2) = "«·«Ã„«·Ì"
    .TextMatrix(0, .Cols - 1) = "«·„” Õﬁ"
End If

.ColWidth(0) = 2000
.ColWidth(1) = 1300
.ColWidth(2) = 1000
.ColWidth(3) = 1000

For i = 3 To .Cols - 2
    .ColWidth(i) = 1000
Next
.ColWidth(.Cols - 1) = 1200
.MergeCells = flexMergeFree
For i = 2 To .Cols - 1
    .ColDataType(i) = flexDTDouble
Next

If .Cols > 8 Then
    Dim nTotal_Charge As Double, nRow As Long
    For nRow = 1 To .Rows - 1
        nTotal_Charge = 0
        For i = 8 To .Cols - 3
            nTotal_Charge = nTotal_Charge + Val(.TextMatrix(nRow, i))
        Next
        .TextMatrix(nRow, .Cols - 2) = nTotal_Charge
        .TextMatrix(nRow, .Cols - 1) = nTotal_Charge - Val(.TextMatrix(nRow, 6))
    Next
End If

.SubtotalPosition = flexSTBelow
For i = 7 To .Cols - 1
    .Subtotal flexSTSum, 0, i, "#0.00", vbYellow, , True, "≈Ã„«·Ì " & "%s"
    .Subtotal flexSTSum, -1, i, "#0.00", vbGreen, , True, "«·≈Ã„«·Ì "
Next

For i = 1 To grid1.Rows - 1
    If grid1.TextMatrix(i, 0) = "≈Ã„«·Ì " & "%s" Then
        grid1.TextMatrix(i, 0) = "≈Ã„«·Ì »œÊ‰  ”ÊÌ…"
    End If
Next
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
.SubtotalPosition = flexSTAbove

StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
'If .Rows > 1 Then
'    .TextMatrix(1, 0) = "«·≈Ã„«·Ì"
'    .TextMatrix(1, 1) = "«·≈Ã„«·Ì"
'    .MergeRow(1) = True
'End If
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

Private Sub xBal_Done_Click()

End Sub

Private Sub xDriver_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    DriverLookupAll Me, oSearchDriver, "DRIVER = 1"
End If
End Sub
Private Sub xDriver_LostFocus()
If Not xDriver.MatchedWithList Then
    xDriver.BoundText = RetZero(xDriver.Text)
    If Not xDriver.MatchedWithList Then xDriver.BoundText = ""
End If
End Sub
Private Sub xTrust_KeyUp(KeyCode As Integer, Shift As Integer)
If Not xDriver.MatchedWithList Then Exit Sub
If KeyCode = 112 Then Trust_LookupAll Me, oSearch, "trust_h.BOX = " & MyParn(xDriver.BoundText)
End Sub
Sub myProc()
If ActiveControl.Name = xTrust.Name Then
    xTrust.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
Else
    xDriver.BoundText = oSearchDriver.grid1.TextMatrix(oSearchDriver.grid1.Row, 0)
    Unload oSearchDriver
End If
End Sub
Private Sub xTrust_LostFocus()
If Trim(xTrust.Text) <> "" Then xTrust.Text = RetZero(xTrust.Text)
End Sub
Private Sub xTrust_Validate(Cancel As Boolean)
If xTrust.Text <> "" Then
    Dim cString As String
    cString = "Select doc_no from trust_h where doc_no = " & RetZero(xTrust.Text)
    If xDriver.MatchedWithList Then cString = cString & turn(cString) & "box = " & MyParn(xDriver.BoundText)
    If IsEmpty(GetField(cString)) Then
        Inform "·«  ÊÃœ  ”ÊÌ… »Â–« «·—ﬁ„ ·Â–« «·”«∆ﬁ"
    End If
End If
End Sub
