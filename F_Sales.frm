VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form F_SALES 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8490
   ClientLeft      =   2430
   ClientTop       =   705
   ClientWidth     =   11400
   FillColor       =   &H00D3BD78&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSComCtl2.MonthView XDATE 
      Height          =   2370
      Left            =   11175
      TabIndex        =   0
      Top             =   75
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   52297735
      CurrentDate     =   39995
   End
   Begin VSFlex7Ctl.VSFlexGrid VsGr 
      Bindings        =   "F_Sales.frx":0000
      Height          =   3660
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   10740
      _cx             =   18944
      _cy             =   6456
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14220542
      ForeColorSel    =   64
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   12975
      Top             =   3750
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
End
Attribute VB_Name = "F_SALES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set Grid1.DataSource = data1
    data1.ConnectionString = CON.ConnectionString
    
'    FixGrid
    VsGr.Rows = 1

End Sub
Private Sub XDATE_DateClick(ByVal DateClicked As Date)
   

cWhere = " [date] = " & DateSq(XDATE.Value) & "  TYPE = '6' "
cField1 = myiif(cWhere, "[OUT]") & " AS Q_SALES"

cWhere = " [date] = " & DateSq(XDATE.Value) & "  TYPE = '3' "
cField2 = myiif(cWhere, "[IN]") & " AS Q_RET "

cWhere = " [date] = " & DateSq(XDATE.Value) & "  TYPE = '3' OR TYPE = '6'  "
cField3 = myiif(cWhere, "[OUT] - [IN]") & " AS Q_RET "


cWhere = " [date] = " & DateSq(XDATE.Value) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField4 = myiif(cWhere, "Val(( FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_11.PRICE & '')*(1-(Val(FILE1_11.DISCOUNT & '')/100))") & " AS TV_SALES"

cWhere = " [date] = " & DateSq(XDATE.Value) & " AND ( TYPE = '6' OR TYPE = '3' )"
cField5 = myiif(cWhere, "Val((FILE1_11.OUT - FILE1_11.[IN] ) & '')* Val(FILE1_10.PRICE & '') ") & " AS TV_PRICE"

With Grid1
'                           0               1                 2                3
    cStrAll = "  SELECT file1_10sc.code as c_sec, file1_10sc.desca  as secdesca ,  FILE1_50G.code as mgrcode , FILE1_50G.DESCA as mgrdesca ,  FILE1_50.code as grcode , FILE1_50.DESCA as grdesc ,  " & _
                cField1 & " , " & cField2 & " , " & cField3 & " , " & cField4 & " , " & cField5 & _
                " FROM (((FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) INNER JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] =  FILE1_10SC.CODE) LEFT JOIN FILE1_50G ON FILE1_50.GROUP = FILE1_50G.CODE  WHERE TRUE "
    cStrAll = cStrAll & " GROUP BY FILE1_50.code, FILE1_50.DESCA, FILE1_50G.code, FILE1_50G.DESCA, FILE1_10SC.CODE, FILE1_10SC.DESCA ORDER BY FILE1_10SC.DESCA , FILE1_50G.DESCA , FILE1_50.DESCA "
    data1.RecordSource = cStrAll
    data1.Refresh
End With
End Sub
