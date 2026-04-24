VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form VsSupp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "√—’œ… «·„Ê—œÌ‰"
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
   Begin VB.TextBox Xcode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9300
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   675
      Width           =   1515
   End
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
      TabIndex        =   0
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
      TabIndex        =   1
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
      Left            =   90
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   525
      Width           =   3000
   End
   Begin VB.CheckBox xMove 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ê—œÌ‰ ·Â„  ⁄«„·"
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
      TabIndex        =   11
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
      TabIndex        =   4
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
      TabIndex        =   6
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid InvGrid1 
      Height          =   7020
      Left            =   75
      TabIndex        =   14
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
      FormatString    =   $"VsSupp.frx":0000
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
      Bindings        =   "VsSupp.frx":0088
      DataSource      =   "Data1"
      Height          =   330
      Left            =   1740
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "≈”„"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   675
      Width           =   315
   End
   Begin VB.Label xDesca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5775
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   675
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   10950
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   330
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
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   255
      Visible         =   0   'False
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   255
      Width           =   780
   End
   Begin VB.Label x3 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x2 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   8250
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label x4 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   975
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   8250
      Width           =   240
   End
End
Attribute VB_Name = "VsSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClientTable As Recordset
Dim Bal1Table As Recordset
Dim Bal2Table As Recordset
Dim Bal3Table As Recordset
Dim ChqTable As Recordset
Dim CashTable As Recordset
Dim Cash0Table As Recordset
Dim PayChqTable As Recordset
Dim PayChq2Table As Recordset
Dim LastSalTable As Recordset
Dim LastCashTable As Recordset
Dim LastChqTable As Recordset
Private Sub Cmd_Print_Click()
    Dim cHead As String
    Dim cHead2 As String
    cHead2 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï " & xDate2.Text
    cHead = "»Ì«‰ »√—’œ… Ê  ⁄«„·«  «·„Ê—œÌ‰  "
    Load PrintGrd
    PrintGrd.doprint InvGrid1, 0.9, -2, cHead, cHead2, , False, True, 8
    PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set VsClient = Nothing
End Sub
Private Sub CmdOk1_Click()
'If Not IsDate(xdate1.Text) Or Not IsDate(xDate2.Text) Then
'    MsgBox " ”ÃÌ· «· «—ÌŒ"
'    Exit Sub
'End If
Dim datatable As Recordset
Dim ssaltable As Recordset
Dim nBal1 As Double, nBal2 As Double, nBal3 As Double
Dim nTot As Double
Dim nSal, nPay, nFBal As Double
Dim tdisc1lTable As Recordset
Dim tdisc2lTable As Recordset

cStr1 = " SELECT ITEMCODE.LastOfCODE, Sum(FILE1_11.OUT) AS TOUT, Sum(FILE1_11.[IN]) AS  TIN , Sum(FILE1_11.OUT * ITEMCODE.price ) AS VOUT , Sum(FILE1_11.[IN] * ITEMCODE.price ) AS  VIN " & _
        " FROM FILE1_11 LEFT JOIN ITEMCODE ON FILE1_11.ITEM = ITEMCODE.ITEM WHERE FILE1_11.ITEM IS NOT NULL "
cStr1 = cStr1 & " GROUP BY ITEMCODE.LastOfCODE "
Set datatable = mydb.OpenRecordset(cStr1)

cStr1 = " SELECT ITEMCODE.LastOfCODE, Sum(FILE6_20.TOTAL) AS TTOTAL, Sum(FILE6_20.QUANT) AS  TQUANT " & _
        " FROM FILE6_20 LEFT JOIN ITEMCODE ON FILE6_20.ITEM = ITEMCODE.ITEM WHERE FILE6_20.DOC_NO IS NOT NULL "
If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND FILE6_20.DATE >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cStr1 = cStr1 & " AND FILE6_20.DATE <= " & DateSql(xDate2.Text)
cStr1 = cStr1 & " GROUP BY ITEMCODE.LastOfCODE "
Set ssaltable = mydb.OpenRecordset(cStr1)

CFIELD2 = myiif("FILE4_11!TYPE = '4' ", "[sal]") & " as T_sal ,"
cField3 = myiif("FILE4_11!TYPE = '5' ", "[pay]") & " as T_rSal, "
cField7 = myiif("FILE4_11!TYPE = '7' ", "[pay]") & " as T_cash, "
cField0 = myiif("FILE4_11!TYPE = '0' ", "[pay]") & " as T_cash0 "


cString = " SELECT FILE4_10.CODE, FILE4_10.DESCA, file4_10.group ,  " & _
           CFIELD2 & cField3 & cField7 & cField0 & _
          " FROM FILE4_10 LEFT JOIN FILE4_11 ON FILE4_10.CODE = FILE4_11.CODE where file4_10.code is not null "
If Xcode.Text <> "" Then cString = cString & " AND FILE4_10.code = " & MyParn(Xcode.Text)
If IsDate(xDate1.Text) Then cString = cString & " AND FILE4_11.DATE >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " AND FILE4_11.DATE <= " & DateSql(xDate2.Text)
cString = cString & " GROUP BY FILE4_10.CODE ,FILE4_10.group, FILE4_10.DESCA "
Set ClientTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT Sum(FILE4_11.SAL) AS TSAL, Sum(FILE4_11.PAY) AS TPAY  , FILE4_11.CODE " & _
          " FROM FILE4_11 where  ( file4_11.SHOW = '1' OR file4_11.SHOW = '2' ) "
If IsDate(xDate1.Text) Then cString = cString & "  and FILE4_11.DATE < " & DateSql(xDate1.Text)
cString = cString & " GROUP BY FILE4_11.CODE  "
Set Bal1Table = mydb.OpenRecordset(cString, dbOpenDynaset)


cString = " SELECT Sum(FILE4_11.SAL) AS TSAL, Sum(FILE4_11.PAY) AS TPAY  , FILE4_11.CODE  " & _
          " FROM FILE4_11 where ( file4_11.SHOW = '1' OR file4_11.SHOW = '2' ) "
If IsDate(xDate2.Text) Then cString = cString & " AND FILE4_11.DATE <= " & DateSql(xDate2.Text)
           cString = cString & " GROUP BY FILE4_11.CODE "
Set Bal2Table = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT Sum(FILE4_11.SAL) AS TSAL, Sum(FILE4_11.PAY) AS TPAY  , FILE4_11.CODE  " & _
          " FROM FILE4_11 where ( file4_11.SHOW = '1' OR file4_11.SHOW = '3' ) "
If IsDate(xDate2.Text) Then cString = cString & " AND FILE4_11.DATE <= " & DateSql(xDate2.Text)
cString = cString & " GROUP BY FILE4_11.CODE  "
Set Bal3Table = mydb.OpenRecordset(cString, dbOpenDynaset)


cStr1 = " SELECT Sum(file5_23.VALUE) AS TVALUE, FILE5_21.CODE FROM FILE5_21 RIGHT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no  WHERE FILE5_21.CODE IS NOT NULL "
If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND FILE5_23.DATE >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cStr1 = cStr1 & " AND FILE5_23.DATE <= " & DateSql(xDate2.Text)
cStr1 = cStr1 & " GROUP BY FILE5_21.CODE "
Set PayChqTable = mydb.OpenRecordset(cStr1, dbOpenDynaset)

cStr1 = " SELECT FILE5_21.CODE, Sum(FILE5_21.VALUE) AS TVALUE FROM FILE5_21 GROUP BY FILE5_21.CODE "
Set ChqTable = mydb.OpenRecordset(cStr1, dbOpenDynaset)

cStr1 = " SELECT Sum(file5_23.VALUE) AS TPAY, FILE5_21.CODE FROM FILE5_21 LEFT JOIN file5_23 ON FILE5_21.SER_NO = file5_23.ser_no GROUP BY FILE5_21.CODE "
Set Chq2Table = mydb.OpenRecordset(cStr1, dbOpenDynaset)


cString = " SELECT FILE7_20.CODE, Max(FILE7_20.DATE) AS LDATE FROM FILE7_20 GROUP BY FILE7_20.CODE "
Set LastSalTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE8_40.CODE, Max(FILE8_40.DATE) AS LDATE FROM FILE8_40 GROUP BY FILE8_40.CODE "
Set LastCashTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE5_23.CODE, Max(FILE5_23.DATE) AS LDATE FROM FILE5_23 GROUP BY FILE5_23.CODE "
Set LastChqTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE7_20.CODE, sum(FILE7_20.total) AS tdisc FROM FILE7_20 where store = 'zz' GROUP BY FILE7_20.CODE "
Set tdisc1lTable = mydb.OpenRecordset(cString, dbOpenDynaset)

cString = " SELECT FILE6_11.CODE, sum(FILE6_11.total) AS tdisc FROM FILE6_11 where store = 'zz'  GROUP BY FILE6_11.CODE "
Set tdisc2lTable = mydb.OpenRecordset(cString, dbOpenDynaset)


With InvGrid1
    i = 1
    InvGrid1.Rows = 1
    nBal1 = 0
    nBal2 = 0
    nBal3 = 0
    If ClientTable.RecordCount = 0 Then Exit Sub
    ClientTable.MoveFirst
    Do While True
        .AddItem ""
        .TextMatrix(i, 0) = ClientTable.CODE
        .TextMatrix(i, 1) = TurnValue(ClientTable.Desca, Null, "")
        
        nBal1 = 0
        nBal2 = 0
        nBal3 = 0
        Bal1Table.FindFirst " CODE = " & MyParn(.TextMatrix(i, 0))
        If Not Bal1Table.NoMatch Then
            nBal1 = TurnValue(Bal1Table.TSAL, Null, 0) - TurnValue(Bal1Table.tpay, Null, 0)
        End If
        
        Bal2Table.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not Bal2Table.NoMatch Then nBal2 = TurnValue(Bal2Table.TSAL, Null, 0) - TurnValue(Bal2Table.tpay, Null, 0)
        
        Bal3Table.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not Bal3Table.NoMatch Then
            nBal3 = TurnValue(Bal3Table.TSAL, Null, 0) - TurnValue(Bal3Table.tpay, Null, 0)
        End If
        
        .TextMatrix(i, 2) = Format(nBal1, "#0.00")
        .TextMatrix(i, 10) = Format(nBal2, "#0.00")
        .TextMatrix(i, 11) = Format(nBal3, "#0.00")

        .TextMatrix(i, 3) = TurnValue(ClientTable.T_SAL, Null, "")
        .TextMatrix(i, 4) = TurnValue(ClientTable.T_Rsal, Null, "")
        
'       .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i, 4)), "#0.00")
        tdisc1lTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not tdisc1lTable.NoMatch Then .TextMatrix(i, 5) = Format(tdisc1lTable.Tdisc, "#0.00")
    
        tdisc2lTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not tdisc2lTable.NoMatch Then .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 5)) - TurnValue(tdisc2lTable.Tdisc, Null, 0), "#0.00")
        .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 5)) * -1, "0#.00")
        .TextMatrix(i, 6) = TurnValue(ClientTable.T_cash, Null, "")
        
        PayChqTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not PayChqTable.NoMatch Then .TextMatrix(i, 7) = Format(PayChqTable.TValue, "#0.00")
        
        .TextMatrix(i, 8) = TurnValue(ClientTable.T_cash0, Null, "")
        
        ChqTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not ChqTable.NoMatch Then
            Chq2Table.FindFirst " CODE = " & MyParn(ClientTable.CODE)
            If Not Chq2Table.NoMatch Then
                .TextMatrix(i, 9) = Format(TurnValue(ChqTable.TValue, Null, 0) - TurnValue(Chq2Table.tpay, Null, 0), "#0.00")
            Else
                .TextMatrix(i, 9) = Format(TurnValue(ChqTable.TValue, Null, 0), "#0.00")
            End If
        End If
        LastSalTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
        If Not LastSalTable.NoMatch Then .TextMatrix(i, 12) = Format(LastSalTable.ldate, "DD-MM-YYYY")
        
        datatable.FindFirst " LastOfCODE = " & MyParn(ClientTable.CODE)
        If Not datatable.NoMatch Then .TextMatrix(i, 13) = Format(TurnValue(datatable.VIN, Null, 0) - TurnValue(datatable.VOUT, Null, 0), "#0.00")
        
        ssaltable.FindFirst " LastOfCODE = " & MyParn(ClientTable.CODE)
        If Not ssaltable.NoMatch Then .TextMatrix(i, 12) = Format(TurnValue(ssaltable.TTOTAL, Null, 0), "#0.00")
        
'        LastCashTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
'        If Not LastCashTable.NoMatch Then .TextMatrix(i, 13) = Format(LastCashTable.ldate, "DD-MM-YYYY")
'
'        LastChqTable.FindFirst " CODE = " & MyParn(ClientTable.CODE)
'        If Not LastChqTable.NoMatch Then
'            If IsDate(.TextMatrix(i, 13)) Then
'                If DateValue(.TextMatrix(i, 13)) < LastChqTable.ldate Then .TextMatrix(i, 13) = Format(LastChqTable.ldate, "DD-MM-YYYY")
'            End If
'        End If
        
        ClientTable.MoveNext
        If ClientTable.EOF Then Exit Do
        i = i + 1
    Loop
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 3, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 4, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 5, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 6, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 7, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 8, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 9, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 10, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 11, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 12, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 13, "#0", , RGB(255, 0, 0), True, " "
End With
Me.MousePointer = 0
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Set CustTable = mydb.OpenRecordset("FILE4_10", dbOpenDynaset)
Set SuppTable = mydb.OpenRecordset("FILE4_10", dbOpenDynaset)
Set ItemTable = mydb.OpenRecordset("FILE1_10", dbOpenDynaset)
Data1.DatabaseName = MdbPath
Data1.RecordSource = "SELECT * FROM FILE1_70 WHERE FLAG = 3 "
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
xDate1.Text = "1-4-2006"
xDate2.Text = Date
With InvGrid1
    .ExplorerBar = flexExSortShow
    .Editable = flexEDNone
    .Cols = 14
    .Rows = 1
    
    .FixedRows = 1
    .FrozenCols = 2
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next i
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1100
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1100
    .ColWidth(11) = 1100
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    
    .RowHeight(0) = 700
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "≈”„"
    .TextMatrix(0, 2) = "—’Ìœ √Ê·"
    .TextMatrix(0, 3) = "„‘ —Ì« "
    .TextMatrix(0, 4) = "„— Ã⁄« "
    .TextMatrix(0, 5) = "Œ’„ «·›Ê« Ì—"
    .TextMatrix(0, 6) = "œ›⁄«  ‰ﬁœÏ"
    .TextMatrix(0, 7) = "”œ«œ ‘Ìﬂ« "
    .TextMatrix(0, 8) = " ”ÊÌ…"
    .TextMatrix(0, 9) = "‘Ìﬂ«  €Ì— „”œœ… "
    .TextMatrix(0, 10) = "—’Ìœ œ› —Ï"
    .TextMatrix(0, 11) = "—’Ìœ ‰ﬁœÏ"
    .TextMatrix(0, 12) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 13) = "—’Ìœ √’‰«›"
    
    For i = 2 To 11
        .ColDataType(i) = flexDTDouble
'        .ColFormat(I) = "#0.00"
    Next i
    .ColDataType(12) = flexDTDate
    .ColDataType(13) = flexDTDate
End With
End Sub
Private Sub InvGrid1_EnterCell()
    With InvGrid1
    If .row > 2 Then
'        If .Col = .Cols - 1 Then .Editable = flexEDKbdMouse
'        If .Col <> .Cols - 1 Then .Editable = flexEDNone
    End If
    End With
End Sub
Private Sub InvGrid1_DBLClick()
With InvGrid1
If Val(.TextMatrix(.row, .Col)) <> 0 Then
    Select Case .Col
        Case 3, 4, 6, 7, 8, 9
            ViewS_S.Show 1
    End Select
End If
End With
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê From File4_10"
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.grid1, 0)
xDesca.Caption = GrdText(Search.grid1, 1)
Unload Search
End Sub

