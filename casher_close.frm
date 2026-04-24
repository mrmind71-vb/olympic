VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form casher_closefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdChangedate 
      Caption         =   "› Õ ÌÊ„ ÃœÌœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2340
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Õ›Ÿ"
      Top             =   1125
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   465
      Left            =   2340
      MaskColor       =   &H00FFFFFF&
      Picture         =   "casher_close.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Õ›Ÿ"
      Top             =   1620
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CheckBox chkPrint 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·€«¡ «·ÿ»«⁄…"
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
      Left            =   945
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1755
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   7395
      Begin VB.TextBox xValue 
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
         Height          =   345
         Left            =   3600
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   1260
         Width           =   2400
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   345
         Left            =   2655
         TabIndex        =   16
         Top             =   900
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1620
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„’«—Ì› :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1665
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„»·€ «·„”·„ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Œ“‰… «·«œ«—… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   990
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ«‘Ì— :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   720
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   540
         Width           =   5820
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   675
      End
      Begin VB.Label xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1725
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   4965
      Visible         =   0   'False
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   90
      Top             =   315
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Height          =   2445
      Left            =   90
      TabIndex        =   1
      Top             =   2115
      Width           =   10950
      _cx             =   19315
      _cy             =   4313
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
      Cols            =   5
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4545
      Width           =   1725
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "casher_close.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
   End
End
Attribute VB_Name = "casher_closefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim oSearchCode As New Search3
Private Sub myloadgrd()
Dim cString As String, cWhere As String
cField = myiif("FLAG = 1 AND CLOSED = 1", "AMOUNT") & " AS Cash"
cField = cField & turn(cField, ",") & _
        myiif("FLAG = 2", "AMOUNT") & " AS Charge"
cField = cField & turn(cField, ",") & _
        "'' AS CASH_NET"
cField = cField & turn(cField, ",") & _
        myiif("FLAG = 3", "AMOUNT") & " as Trans"
cString = "SELECT " & cField & _
          " FROM CASHER_CLOSE"
cString = cString & turn(cString) & "DATE = " & DateSq(dSalesDate)
cString = cString & turn(cString) & "BOX = " & MyParn(sboxSales)
cString = cString & turn(cString, " HAVING ", " AND ") & myiif("FLAG = 3", "AMOUNT") & " > 0"
data11.RecordSource = cString
data11.Refresh
Fixgrd
End Sub
Private Sub cmdChangedate_Click()
If MsgBox("› Õ ÌÊ„ ÃœÌœ", vbOKCancel + vbDefaultButton2) = vbOK Then
    cString = "UPDATE datesales SET DATE = " & addDate(Format(DateAdd("d", 1, dSalesDate), "DD-MM-YYYY"))
    con.BeginTrans
    On Error GoTo myerror
    con.Execute cString
    con.CommitTrans
    dSalesDate = Format(DateAdd("d", 1, dSalesDate), "DD-MM-YYYY")
    Firsttitle = Secondtitle & Format(dSalesDate, "DD-MM-YYYY")
    main.Caption = Firsttitle
    salesfrm.Caption = Format(dSalesDate, "DD-MM-YYYY")
    salesfrm.mydefine
    Inform " „  €ÌÌ— «·ÌÊ„ »‰Ã«Õ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub


Private Sub CmdExit_Click()
'If myreplace Then
'    If Me.chkPrint.Value = 0 Then
'        If Not doprint Then
'            Inform "·„ Ì „ﬂ «·‰Ÿ«„ „‰ «·ÿ»«⁄…!! ·„ Ì „ «·Õ›Ÿ"
'            Exit Sub
'        End If
'    End If
'    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
'    Unload Me
'Else
'    If MsgBox("·„ Ì „ «·Õ›Ÿ !!  —«Ã⁄ ⁄‰ «·Œ—ÊÃ", vbOKCancel) <> vbOK Then Unload Me
'End If
Unload Me
End Sub
Private Sub cmdSave_Click()
If HandleOpen Then Inform " „ Õ–› «·»Ê‰«  «·„› ÊÕ… »‰Ã«Õ"
'If Not CheckOpen Then Exit Sub
If myreplace Then
    xValue.Text = ""
    cmdSave.Enabled = False
    myloadgrd
End If
End Sub

Private Sub Command1_Click()
HandleOpen
End Sub

Private Sub Form_Activate()
If sboxSales = "" Then
    MsgBox "·«  ÊÃœ Œ“‰… „»Ì⁄« "
    Unload Me
'ElseIf Not xBox.MatchedWithList Then
'    MsgBox "·«  ÊÃœ Œ“‰… ≈œ«—…"
'    Unload Me
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
addSetting "print", chkPrint.Value, TempSave(Me)
closeCon con
Set charge_Cashfrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then GrdDesc Row
If validRow(Row) And Row = grid1.Rows - 1 Then myAddItem
CalcTotals
End Sub
Private Sub CalcTotals()
Dim nTotal As Long
For i = 1 To grid1.Rows - 1
    nTotal = Val(grid1.TextMatrix(i, 3)) + nTotal
Next
xtotal.Caption = Myvalue(nTotal)
End Sub
Private Sub grid1_EnterCell()
If bedit Then
    grid1.Editable = IIf(grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Or grid1.Col = 1, flexEDNone, flexEDKbdMouse)
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) = "" Then
    If MsgBox("Õ–› !! Â· «‰  „Ê«›ﬁ ø", vbYesNo) = vbYes Then
        '        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        '            con.BeginTrans
        '            On Error GoTo myerror
        '            Dim cString As String
        '            cString = "Delete from resale_codes"
        '            cString = cString & turn(cString) & " ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        '            con.Execute cString
        '            con.CommitTrans
        '        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Form_Load()
xdate.Caption = Format(dSalesDate, "DD-MM-YYYY")
If sboxSales <> "" Then xdesca.Caption = GetDesca("Select desca from file0_50 where code = " & MyParn(sboxSales))

chkPrint.Value = Val(RetSetting("print", TempSave(Me)))
bedit = True
openCon con

DATA1.ConnectionString = strCon
DATA1.RecordSource = "SELECT * FROM FILE0_50"
Set xBox.RowSource = DATA1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
xBox.BoundText = "00"
'If Not xBox.MatchedWithList Then xBox.BoundText = ""

Set grid1.DataSource = data11
data11.ConnectionString = strCon
myloadgrd
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then Cancel = True
ElseIf Col = 3 Then
   If Not IsNumeric(grid1.EditText) Then
        MsgBox "Numeric Value Requiered"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, grid1.Cols - 1) = "" Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 3)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub Fixgrd()
Dim Col As Long, Row As Long
With grid1
.Cols = 5
.ColWidth(0) = 1500
.ColWidth(1) = 1500
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.TextMatrix(0, 0) = "‰ﬁœÌ… »Ì⁄"
.TextMatrix(0, 1) = "«·„’—Ê›"
.TextMatrix(0, 2) = "„»·€ „› —÷"
.TextMatrix(0, 3) = "„»·€ „”·„"
.TextMatrix(0, 4) = "«·Õ«·…"
For Col = 0 To grid1.Cols - 1
    .ColAlignment(Col) = flexAlignRightCenter
Next
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 2) = Val(.TextMatrix(i, 0)) - Val(.TextMatrix(i, 1))
    grid1.TextMatrix(i, 4) = Myvalue(Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 3)))
    If Val(grid1.TextMatrix(i, 4)) = 0 Then
        grid1.TextMatrix(i, 4) = "·« ÌÊÃœ ⁄Ã“"
    Else
        grid1.TextMatrix(i, 4) = Abs(Val(grid1.TextMatrix(i, 4))) & turn(grid1.TextMatrix(i, 4), " ") & IIf(Val(grid1.TextMatrix(i, 4)) < 0, "“Ì«œ…", "⁄Ã“")
    End If
Next
End With
End Sub
Private Sub xStatus_Change()
myloadgrd
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    grdLookup
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1 + IIf(Col = 0, 1, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, 0
    grid1.ShowCell Row + 1, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Function myreplace() As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(RetZero(Val(Newflag("FILE0_51", "CODE")), 3))

aInsert(2, 0) = "Date"
aInsert(2, 1) = addDate(dSalesDate)

aInsert(3, 0) = "NO1"
aInsert(3, 1) = addstring(sboxSales)

aInsert(4, 0) = "NO2"
aInsert(4, 1) = addstring(xBox.BoundText)

aInsert(5, 0) = "[VALUE]"
aInsert(5, 1) = Val(xValue.Text)

con.BeginTrans
On Error GoTo myerror
Dim cString As String
cString = "UPDATE FILE6_20H SET FILE6_20H.CLOSED = 1"
cString = cString & turn(cString) & "FILE6_20H.DATE = " & DateSq(dSalesDate)
cString = cString & turn(cString) & "FILE6_20H.BOX = " & DateSq(sboxSales)
con.Execute cString
con.Execute CreateInsert(aInsert, "FILE0_51")
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 1) = ""
If Trim(grid1.TextMatrix(Row, 0)) = "" Then Exit Sub
grid1.TextMatrix(Row, 0) = RetZero(grid1.TextMatrix(Row, 0), 3)
grid1.TextMatrix(Row, 1) = GetDesca("Select Desca From file8_51 Where code = " & MyParn(grid1.TextMatrix(Row, 0))) & ""
End Sub
Sub myProc()
grid1.TextMatrix(grid1.Row, 0) = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
grid1.TextMatrix(grid1.Row, 1) = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 1)
grid1_AfterEdit grid1.Row, grid1.Col
CellPos 13, grid1.Row, grid1.Col
Unload oSearchCode
End Sub
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE8_51"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCode.Caption = "≈” ⁄·«„ "
oSearchCode.Show 1
End Sub
Private Sub myAddItem()
grid1.AddItem ""
End Sub
Private Function doprint() As Boolean
Dim nTotal As Double, NROWS As Long
On Error GoTo myerror
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
        nTotal = nTotal + Val(grid1.TextMatrix(i, 3))
        NROWS = NROWS + 1
    End If
Next

If NROWS = 0 Then
    doprint = True
    Exit Function
End If

With grid1
For i = 1 To grid1.Rows - 2
    If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
        temptable.AddNew
        temptable!str1 = "«ﬁ—«— «” ·«„ ‰ﬁœÌ…"
        temptable!str2 = ArbString("«· «—ÌŒ : " & Format(dSalesDate, "yyyy/mm/dd"))
        temptable!str3 = ArbString("√ﬁ— «‰« «·”Ìœ/.................................................")
        temptable!str4 = "»√‰‰Ì «” ·„  „»·€ Êﬁœ—… : " & nTotal & " Ã‰ÌÂ"
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & MyOnly(nTotal)
        temptable!str4 = temptable!str4 & turn(temptable!str4 & " ", " ") & "‰ŸÌ— :"
        temptable!str4 = ArbString(temptable!str4)

        'temptable!str6 = TurnValue(ArbString(.Cell(flexcpTextDisplay, i, 0, i, 0) & turn(.TextMatrix(i, 2), "[" & .TextMatrix(i, 2) & "]")))
        temptable!str6 = TurnValue(.TextMatrix(i, 2))
        temptable!val2 = Val(.TextMatrix(i, 3))
        temptable.Update
    End If
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    doprint = True
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
REPORT1.Reset
FixPrinter REPORT1
REPORT1.ReportFileName = App.Path & "\reports\chargepaid.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.Destination = crptToPrinter
REPORT1.Action = 1
closeCon:
temptable.Close
Set temptable = Nothing
doprint = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function
Private Function CheckOpen() As Boolean
Dim cString As String, nCount As Long
cString = "Select count(*) from file6_20h"
cString = cString & turn(cString) & "File6_20h.printed = 0"
cString = cString & turn(cString) & " [DATE] = " & DateSq(dSalesDate)
cString = cString & turn(cString) & "File6_20h.BOX = " & MyParn(sboxSales)
nCount = Val(GetDesca(cString))
If nCount > 0 Then
    MsgBox "Â‰«ﬂ ⁄œœ " & nCount & " »Ê‰«  »Ì⁄ „› ÊÕ…!!«·—Ã«¡ «·Õ–› «Ê «· ”ÃÌ·", vbCritical
    Exit Function
End If
CheckOpen = True
End Function
Private Sub xValue_Change()
cmdSave.Enabled = Val(xValue.Text) <> 0
End Sub
Private Function HandleOpen() As Boolean
con.BeginTrans
On Error GoTo myerror

Dim cStringIn As String, cString As String, cWhere As String, nRecord As Long
cStringIn = "SELECT DOC_NO from file6_20h"
cStringIn = cStringIn & turn(cStringIn) & "File6_20h.printed = 0"

cWhere = "( [DATE] < " & DateSq(dSalesDate) & ")"
cWhere = cWhere & turn(cWhere, " Or ") & _
         "(" & _
            "File6_20h.BOX = " & MyParn(sboxSales) & _
            " AND  DATE = " & DateSq(dSalesDate) & _
          ")"
cWhere = turn(cWhere, "(") & cWhere & turn(cWhere, ")")
If cWhere <> "" Then cStringIn = cStringIn & turn(cStringIn) & cWhere
cString = "DELETE FROM FILE6_20 WHERE FILE6_20.DOC_NO IN (" & _
          cStringIn & _
          ")"

con.Execute cString, nRecord
          
cString = "DELETE from file6_20h"
cString = cString & turn(cString) & "File6_20h.printed = 0"
If cWhere <> "" Then cString = cString & turn(cString) & cWhere
con.Execute cString, nRecord
con.CommitTrans
HandleOpen = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function

