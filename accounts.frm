VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form acountsfrm 
   Caption         =   " ﬁ—Ì— Õ”«» «—»«Õ ÊŒ”«∆—"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Õ”«» «—»«Õ ÊŒ”«∆—"
      TabPicture(0)   =   "accounts.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid1(1)"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Õ”«» «·„ «Ã—…"
      TabPicture(1)   =   "accounts.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grid1(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3255
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   405
         Width           =   9285
         _cx             =   16378
         _cy             =   5741
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
         Cols            =   4
         FixedRows       =   0
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3255
         Index           =   1
         Left            =   -74910
         TabIndex        =   10
         Top             =   405
         Width           =   9285
         _cx             =   16378
         _cy             =   5741
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
         Cols            =   4
         FixedRows       =   0
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
      Begin VB.Frame Frame2 
         Height          =   690
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2970
         Visible         =   0   'False
         Width           =   9285
         Begin VB.Label xTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   4455
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   4695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   690
         Left            =   -74910
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2970
         Visible         =   0   'False
         Width           =   9285
         Begin VB.Label xTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   225
            Width           =   4695
         End
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3870
      Width           =   3840
      Begin VB.CommandButton cmdGo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2520
         Picture         =   "accounts.frx":0038
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         Picture         =   "accounts.frx":252A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1305
         Picture         =   "accounts.frx":4996
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4005
      TabIndex        =   8
      Top             =   3870
      Width           =   5595
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1950
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "«· «—ÌŒ :"
         Height          =   375
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   735
      End
   End
End
Attribute VB_Name = "acountsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
Dim loctable As New ADODB.Recordset, cString As String
Dim nFirstBal As Double, nPurchase As Double, nSales As Double, nLastBal As Double, nTotal0 As Double, nTotal1 As Double, nCharge As Double
If IsDate(xdate1.Text) Then
    sDate = Format(DateAdd("d", -1, xdate1.Text), "YYYY-MM-DD")
    cString = "SELECT SUM( ([IN] - OUT) * dbo.f_item_cost(FILE1_10.ITEM," & MyParn(sDate) & ")) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
    cString = cString & turn(cString) & "DATE < " & DateSq(xdate1.Text)
    nFirstBal = Val(GetDesca(cString))
End If

For I = 0 To 1
    grid1(I).TextMatrix(0, 1) = Format(nFirstBal, "Fixed")
Next

Dim cwhere As String
If IsDate(xdate1.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)

cString = "Select SUM(FILE7_20.TOTAL) FROM FILE7_20  INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nPurchase = nPurchase + Val(GetDesca(cString))


cString = "Select SUM(FILE7_20H.DISCOUNT) FROM FILE7_20H"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nPurchase = nPurchase - Val(GetDesca(cString))

cString = "Select SUM(FILE7_30.TOTAL) FROM FILE7_30  INNER JOIN FILE7_30H ON FILE7_30.DOC_NO = FILE7_30H.DOC_NO"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nPurchase = nPurchase - Val(GetDesca(cString))
'
cString = "Select SUM(FILE7_30H.DISCOUNT) FROM FILE7_30H"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nPurchase = nPurchase + Val(GetDesca(cString))

For I = 0 To 1
    grid1(I).TextMatrix(1, 1) = Format(nPurchase, "Fixed")
Next

cString = "Select SUM(FILE6_20.TOTAL) FROM FILE6_20  INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
cString = cString & turn(cString) & " FILE6_20H.PRINTED = 1"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nSales = Val(GetDesca(cString))
'
cString = "Select SUM(FILE6_20H.DISCOUNT) FROM FILE6_20H"
cString = cString & turn(cString) & " FILE6_20H.PRINTED = 1"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nSales = nSales - Val(GetDesca(cString))
For I = 0 To 1
    grid1(I).TextMatrix(0, 3) = Format(nSales, "Fixed")
Next

If IsDate(XDATE2.Text) Then
    sDate = Format(XDATE2.Text, "YYYY-MM-DD")
    cString = "SELECT SUM( ([IN] - OUT) * dbo.f_item_cost(FILE1_10.ITEM," & MyParn(sDate) & ")) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
    cString = cString & turn(cString) & "DATE <= " & DateSq(XDATE2.Text)
    nLastBal = Val(GetDesca(cString))
End If

For I = 0 To 1
    grid1(I).TextMatrix(1, 3) = Format(nLastBal, "Fixed")
Next
nTotal0 = nSales + nLastBal - nFirstBal - nPurchase

If nTotal0 >= 0 Then
    xTotal(0).Caption = "«—»«Õ : " & Format(nTotal0, "Fixed")
    grid1(0).TextMatrix(grid1(0).Rows - 1, 1) = Format(nTotal0, "Fixed")
Else
    xTotal(0).Caption = "Œ”«∆— : " & Format(Abs(nTotal0), "Fixed")
    grid1(0).TextMatrix(grid1(0).Rows - 1, 3) = Format(Abs(nTotal0), "Fixed")
End If

cString = "Select SUM(FILE8_50.VALUE) FROM FILE8_50  INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO"
If cwhere <> "" Then cString = cString & turn(cString) & cwhere
nCharge = Val(GetDesca(cString))

grid1(1).TextMatrix(2, 3) = Format(nCharge, "Fixed")
nTotal1 = nTotal0 - nCharge

If nTotal1 >= 0 Then
    grid1(1).TextMatrix(grid1(1).Rows - 1, 1) = Format(nTotal1, "Fixed")
    xTotal(1).Caption = "«—»«Õ : " & Format(nTotal1, "Fixed")
Else
    grid1(1).TextMatrix(grid1(1).Rows - 1, 3) = Format(Abs(nTotal1), "Fixed")
    xTotal(0).Caption = "Œ”«∆— : " & Format(Abs(nTotal1), "Fixed")
End If
End Sub

Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    If SSTab1.Tab = 0 Then
        cHead1 = "Õ”«» «—»«Õ ÊŒ”«∆—"
    Else
        cHead1 = "Õ”«» «·„ «Ã—…"
    End If
    If IsDate(xdate1.Text) Then cHead2 = "„‰ : " & Format(xdate1.Text, "YYYY-MM-DD")
    If IsDate(XDATE2.Text) Then cHead2 = cHead2 & turn(cHead2, " ") & "Õ Ì : " & Format(XDATE2.Text, "YYYY-MM-DD")
    If Me.SSTab1.Tab = 1 Then
        PrintGrd.doprint Me.grid1(0), 1, -1, cHead1, cHead2, , False, False, 10, , Array(1)
    Else
        PrintGrd.doprint Me.grid1(1), 1, -1, cHead1, cHead2, , False, False, 10, , Array(1)
    End If
    PrintGrd.Show 1
End Sub

Private Sub Form_Load()
openCon con
Fixgrd
LoadText Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
End Sub

Private Sub xdate1_Change()
cmdGo.Enabled = IsDate(xdate1.Text) And IsDate(XDATE2.Text)
End Sub

Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub

Private Sub xdate2_Change()
cmdGo.Enabled = IsDate(xdate1.Text) And IsDate(XDATE2.Text)
End Sub

Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate XDATE2
End Sub
Private Sub Fixgrd()
For I = 0 To 1
    With grid1(I)
        .Rows = 2
        .ColWidth(0) = 2600
        .ColWidth(1) = 1500
        .ColWidth(2) = 2600
        .ColWidth(3) = 1500
    
        .TextMatrix(0, 0) = "—’Ìœ «Ê· „œ…"
        .TextMatrix(1, 0) = "„‘ —Ì« "
        
        .TextMatrix(0, 2) = "„»Ì⁄« "
        .TextMatrix(1, 2) = "—’Ìœ «Œ— «·„œ…"
    End With
Next
grid1(1).AddItem ""
grid1(1).TextMatrix(2, 2) = "„’«—Ì›"

For I = 0 To 1
    With grid1(I)
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "√—»«Õ"
        .TextMatrix(.Rows - 1, 2) = "Œ”«∆—"
    End With
Next
FixColor
End Sub
Private Sub FixColor()
For I = 0 To 1
    With grid1(I)
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &H8000000F
        .Cell(flexcpBackColor, 0, 2, .Rows - 1, 2) = &H8000000F
    End With
Next
End Sub
