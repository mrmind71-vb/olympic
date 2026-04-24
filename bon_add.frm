VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form bon_addfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇÖÇÝÉ ÈæäÇÊ"
   ClientHeight    =   8655
   ClientLeft      =   15
   ClientTop       =   390
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   10485
   Begin VB.Frame Frame2 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2970
      Width           =   3885
      Begin VB.Label Label7 
         Caption         =   "ÈæäÇÊ ÛíÑ ãÓÊÎÏãÉ"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Width           =   1545
      End
      Begin VB.Label xBonsRest 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1395
         Width           =   2130
      End
      Begin VB.Label Label5 
         Caption         =   "ÇáÈæäÇÊ ÇáãÓÊÎÏãÉ"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label xBonsUsed 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Width           =   2130
      End
      Begin VB.Label Label3 
         Caption         =   "ÚÏÏ ÇáÈæäÇÊ"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label xBons 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "ÑÞã ÇáÏÝÊÑ"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   870
      End
      Begin VB.Label xNoteDesca 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   2130
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ÈæäÇÊ Êã ÇÓÊÎÏÇãåÇ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4860
      Width           =   10230
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   3210
         Left            =   90
         TabIndex        =   3
         Top             =   315
         Width           =   10005
         _cx             =   17648
         _cy             =   5662
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
         BackColorFixed  =   12648384
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
         Cols            =   5
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "ÈæäÇÊ áã ÊÓÊÎÏã ÈÚÏ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   6360
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4425
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   6180
         _cx             =   10901
         _cy             =   7805
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
         Cols            =   2
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
         AutoSearch      =   1
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
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   495
      Top             =   -180
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   2700
      Top             =   0
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
Attribute VB_Name = "bon_addfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nFlag As Long
Public sNote As String, sNoteDesca As String, sBon As String
Public myForm As Form
Dim con As New ADODB.Connection
Private Sub cmdFilter_Click()
myload1
myload2
End Sub
Private Sub cmdExit_Click()
Unload Me
Set travel_addfrm = Nothing
End Sub
Private Sub CmdUndo_Click()
Unload Me
End Sub
Private Sub Form_Load()
nFlag = 1
Me.Caption = IIf(nFlag = 1, "ÈæäÇÊ ÓæáÇÑ-ÈäÒíä", "ÈæäÇÊ ÒíÊ")
openCon con
Set GRID2.DataSource = data12
xNoteDesca.Caption = sNoteDesca
myload2
myload1
xBons.Caption = grid1.Rows + GRID2.Rows - 2
xBonsRest.Caption = grid1.Rows - 1
xBonsUsed.Caption = GRID2.Rows - 1
End Sub
Private Sub myload1()
Dim aRet As Variant, cString As String
aRet = GetFields("select BON,BON_COUNT FROM NOTES_CODES" & IIf(nFlag = 1, "", "_OIL") & " WHERE CODE = " & sNote)
If Not IsEmpty(aRet) Then
    For i = 0 To retFlag(aRet, "BON_COUNT") - 1
         If GRID2.FindRow(Val(retFlag(aRet, "BON")) + i, , 1) = -1 Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
            grid1.TextMatrix(grid1.Rows - 1, 1) = Val(retFlag(aRet, "BON")) + i
         End If
    Next
End If
fixgrd1
End Sub
Private Sub myload2()
With grid1
Dim cString As String
If nFlag = 1 Then
    cString = "SELECT GAS_ORDERS.BON AS [ÑÞã ÇáÈæä], CARS.DESCA AS [ÇáÓíÇÑÉ], CONVERT(VARCHAR(10),GAS_ORDERS.DATE,111) AS [ÇáÊÇÑíÎ], DRIVER.DESCA AS [ÇáÓÇÆÞ]" & _
              " FROM  GAS_ORDERS INNER JOIN CARS ON GAS_ORDERS.CAR = CARS.CODE INNER JOIN DRIVER ON GAS_ORDERS.DRIVER = DRIVER.CODE"
    cString = cString & turn(cString) & "[NOTE] = " & MyParn(sNote)
    cString = cString & " ORDER BY DATE,COUNTER"
Else
    cString = "SELECT OIL_ORDERS.BON AS [ÑÞã ÇáÈæä], CARS.DESCA AS [ÇáÓíÇÑÉ], CONVERT(VARCHAR(10),OIL_ORDERS.DATE,111) AS [ÇáÊÇÑíÎ], DRIVER.DESCA AS [ÇáÓÇÆÞ]" & _
              " FROM  OIL_ORDERS INNER JOIN CARS ON OIL_ORDERS.CAR = CARS.CODE INNER JOIN DRIVER ON OIL_ORDERS.DRIVER = DRIVER.CODE"
    cString = cString & turn(cString) & "[NOTE] = " & MyParn(sNote)
    cString = cString & " ORDER BY DATE,COUNTER"
End If
Set data12.Recordset = myRecordSet(cString, con)
fixgrd2
End With
End Sub
Sub fixgrd2()
With GRID2
'.Cols = 11
.TextMatrix(0, 0) = "ãÓáÓá"
.ColWidth(0) = 800
.ColWidth(1) = 2000
.ColWidth(2) = 2500
.ColWidth(3) = 1300
.ColWidth(4) = 2500
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
    If .TextMatrix(i, 1) = sBon Then
        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbYellow
    End If
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Sub fixgrd1()
With grid1
.FormatString = "ãÓáÓá|ÑÞã ÇáÈæä"
.ColWidth(0) = 1000
.ColWidth(1) = 3000
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set bon_addfrm = Nothing
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xdate2
End Sub

Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate_policy1_GotFocus()
myGotFocus xDate_policy1
End Sub
Private Sub xDate_policy1_LostFocus()
myLostFocus xDate_policy1
myValidDate xDate_policy1
End Sub
Private Sub xDate_Policy2_GotFocus()
myGotFocus xDate_Policy2
End Sub
Private Sub xDate_Policy2_LostFocus()
myLostFocus xDate_Policy2
myValidDate xDate_Policy2
End Sub
Private Sub grid1_dblClick()
If grid1.Row > 0 Then
    myForm.sBon = grid1.TextMatrix(grid1.Row, 1)
    Unload Me
End If
End Sub
