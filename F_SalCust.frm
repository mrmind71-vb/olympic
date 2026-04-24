VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form F_SalCust 
   BackColor       =   &H00E0E0E0&
   Caption         =   "„ «»⁄… „»Ì⁄«   ··⁄„·«¡ „œÌ‰… «·√Õ·«„"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox xVal 
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
      Left            =   5100
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1515
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00E3C7AB&
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
      Height          =   465
      Left            =   4650
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   75
      Width           =   2010
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4950
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   6825
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   75
      Width           =   1515
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
      Left            =   9375
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E x i t"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid VsItem 
      Height          =   5550
      Left            =   150
      TabIndex        =   0
      Top             =   1050
      Width           =   11715
      _cx             =   20664
      _cy             =   9790
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
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
      WallPaperAlignment=   9
   End
   Begin MSDBCtls.DBCombo xStore 
      Bindings        =   "F_SalCust.frx":0000
      DataSource      =   "Data1"
      Height          =   315
      Left            =   8850
      TabIndex        =   6
      Top             =   600
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„»Ì⁄«  √ﬂ»— „‰"
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
      Left            =   7125
      TabIndex        =   10
      Top             =   705
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ „Œ“‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10950
      TabIndex        =   7
      Top             =   600
      Width           =   765
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
      Left            =   8475
      TabIndex        =   5
      Top             =   180
      Width           =   780
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
      Left            =   11025
      TabIndex        =   4
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "F_SalCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As Recordset
Dim TSalTable As Recordset
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdOk_Click()
Dim cStr1 As String
cField1 = myiif("STORE <>  'zz' ", "[TOTAL]") & " as T_SAL ,"
cField2 = myiif("STORE =   'zz' ", "[TOTAL]") & " as T_DISC "

cStr1 = " SELECT FILE3_20.ADDRESS , FILE3_20.CLASS , FILE3_20.SCOOL , FILE6_20.CUST , FILE6_20.DOC_NO , FILE6_20.DATE , FILE3_20.DESCA ,  MSTORE.MSTORE , " & _
        cField1 & cField2 & _
        " FROM (FILE6_20 LEFT JOIN FILE3_20 ON FILE6_20.CUST = FILE3_20.code) LEFT JOIN MSTORE ON FILE6_20.DOC_NO = MSTORE.DOC_NO  WHERE FILE6_20.DOC_NO IS NOT NULL "
If xStore.BoundText <> "" Then cStr1 = cStr1 & " AND MSTORE = " & MyParn(xStore.BoundText)
If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND FILE6_20.DATE >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cStr1 = cStr1 & " AND FILE6_20.DATE <= " & DateSql(xDate2.Text)
cStr1 = cStr1 & " GROUP BY FILE3_20.address, FILE3_20.class, FILE3_20.scool, FILE6_20.CUST, FILE6_20.DOC_NO, FILE6_20.DATE, FILE3_20.DESCA, MSTORE.MSTORE "
Set TSalTable = mydb.OpenRecordset(cStr1)

With VsItem
    .WordWrap = True
    .ExplorerBar = flexExSortShow
    .Rows = 1
    .Cols = 9
    .RowHeight(0) = 700
    .FixedCols = 1
    .FixedRows = 1
    .WordWrap = True
    .ExplorerBar = flexExSortShow
    .TextMatrix(0, 0) = " «—ÌŒ"
    .TextMatrix(0, 1) = "„” ‰œ"
    .TextMatrix(0, 2) = " ·Ì›Ê‰"
    .TextMatrix(0, 3) = "«·≈”„"
    .TextMatrix(0, 4) = "≈Ã„«·Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 5) = "≈Ã„«·Ï ﬁÌ„… Œ’„"
    .TextMatrix(0, 6) = "«·⁄‰Ê«‰"
    .TextMatrix(0, 7) = "«·„œ—”…"
    .TextMatrix(0, 8) = "«·”‰… «·œ—«”Ì…"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 2200
    .ColWidth(7) = 2200
    .ColWidth(8) = 2200
    .FrozenCols = 2

    .Rows = 1
    TSalTable.MoveFirst
    If TSalTable.RecordCount > 0 Then
        Do While Not TSalTable.EOF
            If TSalTable.T_SAL >= Val(xVal.Text) Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = Format(TSalTable.Date, "DD-MM-YYYY")
                .TextMatrix(.Rows - 1, 1) = TSalTable.doc_no
                .TextMatrix(.Rows - 1, 2) = TurnValue(TSalTable.Cust, Null, "")
                .TextMatrix(.Rows - 1, 3) = TurnValue(TSalTable.desca, Null, "")
                .TextMatrix(.Rows - 1, 4) = Format(TSalTable.T_SAL, "#0.00")
                .TextMatrix(.Rows - 1, 5) = Format(TSalTable.T_DISC, "#0.00")
                .TextMatrix(.Rows - 1, 6) = TurnValue(TSalTable.ADDRESS, Null, "")
                .TextMatrix(.Rows - 1, 7) = TurnValue(TSalTable.SCOOL, Null, "")
                .TextMatrix(.Rows - 1, 8) = TurnValue(TSalTable.Class, Null, "")
            End If
            TSalTable.MoveNext
        Loop
    End If
End With
End Sub

Private Sub Form_Load()
data1.DatabaseName = MdbPath
data1.RecordSource = "SELECT CODE , DESCA FROM FILE1_70 WHERE FLAG = 1 "
xStore.ListField = "Desca"
xStore.BoundColumn = "code"
End Sub
