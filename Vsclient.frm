VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form VsClient 
   BackColor       =   &H00E0E0E0&
   Caption         =   "≈Ã„«·Ï  ⁄«„·«  «·⁄„·«¡"
   ClientHeight    =   8490
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9480
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox xDate1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9300
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   150
      Width           =   1515
   End
   Begin VB.TextBox xDate2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   150
      Width           =   1515
   End
   Begin VB.CommandButton CmdOk1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   525
      Width           =   2640
   End
   Begin VB.CheckBox xMove 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄„·«¡ ·Â„  ⁄«„·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   225
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E3C7AB&
      Caption         =   " —«Ã⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1155
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E3C7AB&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid InvGrid1 
      Height          =   7020
      Left            =   75
      TabIndex        =   13
      Top             =   1275
      Width           =   11715
      _cx             =   20664
      _cy             =   12382
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   16761024
      ForeColorSel    =   255
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   2
      GridLinesFixed  =   1
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Vsclient.frx":0000
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Vsclient.frx":0088
      DataSource      =   "Data1"
      Height          =   330
      Left            =   8040
      TabIndex        =   14
      Top             =   630
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      Text            =   "DBCombo1"
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·„Ã„Ê⁄… :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   10890
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   705
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   10950
      TabIndex        =   12
      Top             =   255
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï  «—ÌŒ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8400
      TabIndex        =   11
      Top             =   255
      Width           =   780
   End
   Begin VB.Label x3 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x2 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   8250
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label x4 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   975
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8250
      Width           =   240
   End
End
Attribute VB_Name = "VsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset
Dim CustTable As Recordset
Dim SuppTable As Recordset
Dim itemTable As Recordset
Dim cString As String
Dim SalTable As Recordset
Dim RSalTable As Recordset
Dim DiscSal As Recordset
Dim DiscRSal As Recordset
Dim LastSalTable As Recordset
Dim CashTable As Recordset
Dim CHQTable As Recordset
Dim Cash2Table As Recordset
Dim LastCashTable As Recordset
Dim nView As Byte
Private Sub CMD_PRINT_Click()
    Dim cHead As String
    Dim cHead2 As String
    cHead = firsttitle
    cHead2 = "»Ì«‰ »√—’œ… Ê  ⁄«„·«  «·⁄„·«¡ "
    With InvGrid1
        For I = 1 To .Rows - 1
            If .TextMatrix(I, .Cols - 1) <> "True" Then .RowHidden(I) = True
        Next I
        .ColHidden(.Cols - 1) = True
        .RowHidden(1) = False
    End With
    Load PrintGrd
    PrintGrd.Doprint InvGrid1, 1.3, -2, cHead, cHead2, , False, True, 10
    PrintGrd.Show 1
    With InvGrid1
        For I = 1 To .Rows - 1
            .RowHidden(I) = False
        Next I
        .ColHidden(.Cols - 1) = False
    End With
End Sub
Private Sub CmdExit_Click()
Unload Me
Set VsClient = Nothing
End Sub
Private Sub CmdOk1_Click()
Dim nBal As Double
Dim nTot As Double
Dim nSal, nPay, nFBal As Double
cString = " SELECT FILE3_10.CODE,FILE3_10.f_balance, FILE3_10.DESCA, Sum(FILE3_11.SAL) AS SumSAL, Sum(FILE3_11.PAY) AS SumPAY  " & _
          " FROM FILE3_10 LEFT JOIN FILE3_11 ON FILE3_10.CODE = FILE3_11.CODE where file3_10.code is not null "
If xGroup.BoundText <> "" Then cString = cString & " AND FILE3_10.GROUP = " & MyParn(xGroup.BoundText)
           cString = cString & " GROUP BY FILE3_10.CODE ,FILE3_10.f_balance, FILE3_10.DESCA "
Set ClientTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cStrSal = " SELECT Sum(FILE6_20.TOTAL) AS SumTOTAL, FILE6_20.CODE FROM FILE6_20  WHERE STORE <> 'zz'GROUP BY FILE6_20.CODE  "
Set SalTable = mydb.OpenRecordset(cStrSal, dbOpenDynaset)

cStrSal = " SELECT Sum(FILE6_10.TOTAL) AS SumTOTAL, FILE6_10.CODE FROM FILE6_10 WHERE STORE <> 'zz'GROUP BY FILE6_10.CODE  "
Set RSalTable = mydb.OpenRecordset(cStrSal, dbOpenDynaset)

cStrSal = " SELECT Sum(FILE6_20.TOTAL) AS SumTOTAL, FILE6_20.CODE FROM FILE6_20 WHERE STORE = 'zz'GROUP BY FILE6_20.CODE  "
Set DiscSal = mydb.OpenRecordset(cStrSal, dbOpenDynaset)

cStrSal = " SELECT Sum(FILE6_10.TOTAL) AS SumTOTAL, FILE6_10.CODE FROM FILE6_10 WHERE STORE = 'zz'GROUP BY FILE6_10.CODE  "
Set DiscRSal = mydb.OpenRecordset(cStrSal, dbOpenDynaset)

cString = " SELECT FILE6_20.CODE, Max(FILE6_20.DATE) AS MaxDATE FROM FILE6_20 GROUP BY FILE6_20.CODE "
Set LastSalTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE8_10.CODE, Max(FILE8_10.DATE) AS MaxDATE FROM FILE8_10 GROUP BY FILE8_10.CODE "
Set LastCashTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT CASH , FILE8_10.CODE, Max(FILE8_10.DATE) AS MaxDATE , sum(file8_10.value) as sumtot FROM FILE8_10 GROUP BY FILE8_10.CODE , CASH "
Set CashTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE5_20.CODE, sum(file5_20.value) as sumtot FROM FILE5_20 GROUP BY FILE5_20.CODE "
Set CHQTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT CASH , FILE8_40.CODE , sum(file8_40.value) as sumtot FROM FILE8_40 GROUP BY FILE8_40.CODE , CASH  "
Set Cash2Table = mydb.OpenRecordset(cString, dbOpenDynaset)

With InvGrid1
    ClientTable.MoveFirst
    I = 1
    .Rows = 1
    nTot = 0
    Do While True
        nBal = TurnValue(ClientTable.SUMSAL, Null, 0) - TurnValue(ClientTable.sumpay, Null, 0)
        .AddItem ""
        .TextMatrix(I, .Cols - 1) = True
        .TextMatrix(I, 0) = ClientTable.CODE
        .TextMatrix(I, 1) = TurnValue(ClientTable.DESCA, Null, "")
        .TextMatrix(I, 7) = Format(ClientTable.F_balance, "##0.00")

        CashTable.FindFirst " CASH = TRUE AND CODE = " & MyParn(ClientTable.CODE)
        If Not CashTable.NoMatch Then .TextMatrix(I, 8) = Format(CashTable.SUMTOT, "#0.00")
        Cash2Table.FindFirst " CASH = TRUE AND CODE = " & MyParn(ClientTable.CODE)
        If Not Cash2Table.NoMatch Then .TextMatrix(I, 8) = Format(TurnValue(Val(.TextMatrix(I, 8)) - CashTable.SUMTOT, Null, 0), "#0.00")
        
        CashTable.FindFirst " CASH = FALSE AND CODE = " & MyParn(ClientTable.CODE)
        If Not CashTable.NoMatch Then .TextMatrix(I, 10) = Format(CashTable.SUMTOT, "#0.00")
        Cash2Table.FindFirst " CASH = FALSE AND CODE = " & MyParn(ClientTable.CODE)
        If Not Cash2Table.NoMatch Then .TextMatrix(I, 10) = Format(TurnValue(Val(.TextMatrix(I, 10)) - CashTable.SUMTOT, Null, 0), "#0.00")
        
        CHQTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not CHQTable.NoMatch Then .TextMatrix(I, 9) = Format(CHQTable.SUMTOT, "#0.00")
        
        SalTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not SalTable.NoMatch Then .TextMatrix(I, 2) = Format(SalTable.SUMTOTAL, "#0.00")
        
        RSalTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not RSalTable.NoMatch Then .TextMatrix(I, 3) = Format(RSalTable.SUMTOTAL, "#0.00")
        
        DiscRSal.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not DiscRSal.NoMatch Then .TextMatrix(I, 5) = Format(DiscRSal.SUMTOTAL, "#0.00")
        DiscSal.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not DiscSal.NoMatch Then .TextMatrix(I, 5) = Format(Val(.TextMatrix(I, 5)) - TurnValue(DiscSal.SUMTOTAL, Null, 0), "#0.00")
        
        .TextMatrix(I, 4) = Format(Val(.TextMatrix(I, 2)) - Val(.TextMatrix(I, 3)), "#0.00")
        nTot = nTot + .TextMatrix(I, 4)
        .TextMatrix(I, 6) = Format(Val(.TextMatrix(I, 4)) - Val(.TextMatrix(I, 5)), "#0.00")
        
        .TextMatrix(I, 11) = Format(nBal, "#0.00")
        If Val(.TextMatrix(I, 4)) > 0 Then
            .TextMatrix(I, 12) = Format((Val(.TextMatrix(I, 5)) + Val(.TextMatrix(I, 10))) / .TextMatrix(I, 4) * 100, "#0.00")
        End If
        
        If Val(.TextMatrix(I, 2)) = 0 And Val(.TextMatrix(I, 8)) = 0 And Val(.TextMatrix(I, 7)) = 0 And xMove.Value = 1 Then
            .RemoveItem .Rows - 1
            I = I - 1
        End If
        If ClientTable.EOF Then Exit Do
        ClientTable.MoveNext
        If ClientTable.EOF Then Exit Do
        I = I + 1
    Loop
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 3, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 4, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 10, "##0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 11, "##0", , RGB(255, 0, 0), True, " "
'    .Subtotal flexSTSum, -1, 11, "##0", , RGB(255, 0, 0), True, " "
End With
Me.MousePointer = 0
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Sub myProc()
'If ActiveControl .Name = xSerItem.Name Then
     ActiveControl.Text = GrdText(Search.Grid1, 0)
'End If
Unload Search
End Sub
Private Sub Form_Load()
Set CustTable = mydb.OpenRecordset("FILE3_10", dbOpenDynaset)
Set SuppTable = mydb.OpenRecordset("FILE4_10", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("FILE1_10", dbOpenDynaset)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 11 "
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

With InvGrid1
    .ExplorerBar = flexExSortShow
    .Editable = flexEDNone
    .Cols = 14
    .Rows = 1
    
    .FixedRows = 1
    .FrozenCols = 2
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next I
    .ColWidth(0) = 400
    .ColWidth(1) = 1800
    .ColWidth(2) = 1000
    .ColWidth(3) = 800
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 700
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 800
    .ColWidth(11) = 1000
    .ColWidth(12) = 500
    .ColWidth(13) = 400

    
    .RowHeight(0) = 700
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "≈”„"
    .TextMatrix(0, 2) = "„»Ì⁄« "
    .TextMatrix(0, 3) = "„— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï „»Ì⁄«  „ÊœÌ·« "
    .TextMatrix(0, 5) = "Œ’„ ›Ê« Ì—"
    .TextMatrix(0, 6) = "’«›Ï ›Ê« Ì— »⁄œ «·Œ’„"
    .TextMatrix(0, 7) = "—’Ìœ „—Õ·"
    .TextMatrix(0, 8) = "’«›Ï ”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 9) = "”œ«œ ‘ÌﬂÏ"
    .TextMatrix(0, 10) = "Œ’„ ‰ﬁœÏ"
    .TextMatrix(0, 11) = "—’Ìœ"
    .TextMatrix(0, 12) = "‰”»… ≈Ã„«·Ï Œ’„"
    .TextMatrix(0, 13) = "Print"
    .ColDataType(13) = flexDTBoolean
    For I = 2 To 12
        .ColDataType(I) = flexDTDouble
    Next I
End With
End Sub
Private Sub InvGrid1_DBLClick()
    If InvGrid1.Row > 5 Then
'    x3.Caption = InvGrid1.TextMatrix(InvGrid1.Row, 0)
'    x4.Caption = InvGrid1.TextMatrix(InvGrid1.Row, 1)
'    SubVsC_1.Show 1
    End If
End Sub
Private Sub InvGrid1_EnterCell()
    With InvGrid1
    If .Row > 0 Then
    If .Col = .Cols - 1 Then .Editable = flexEDKbdMouse
    If .Col <> .Cols - 1 Then .Editable = flexEDNone
    End If
    End With
End Sub
