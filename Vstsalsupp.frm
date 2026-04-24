VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form Vstsalsupp 
   Caption         =   "≈Ã„«·Ï ﬁÌ„… „»Ì⁄«  √’‰«› «·„Ê—œÌ‰"
   ClientHeight    =   9165
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12090
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
   ScaleHeight     =   9165
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   2850
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1755
         TabIndex        =   12
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1725
         TabIndex        =   11
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7695
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00E3C7AB&
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E3C7AB&
      Caption         =   " —«Ã⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E3C7AB&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   1050
   End
   Begin VSFlex7LCtl.VSFlexGrid grid1 
      Height          =   7740
      Left            =   90
      TabIndex        =   13
      Top             =   1035
      Width           =   11955
      _cx             =   21087
      _cy             =   13652
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
      SelectionMode   =   1
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
      Editable        =   0
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
   Begin VB.Label x3 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   1275
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x2 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8250
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label x4 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   975
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   8250
      Width           =   240
   End
End
Attribute VB_Name = "Vstsalsupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Print_Click()
    Dim cHead As String
    Dim cHead2 As String
    cHead = Firsttitle
    cHead2 = "»Ì«‰ »⁄œœ Ê ﬁÌ„… „»Ì⁄«  √’‰«› «·„Ê—œÌ‰ "
    Load PrintGrd
    PrintGrd.doprint grid1, 0.9, -2, cHead, cHead2, , False, , 8
    PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set VsClient = Nothing
End Sub
Private Sub CmdOk1_Click()
Dim nBal1 As Double, nBal2 As Double, nBal3 As Double
Dim nTot As Double
Dim nSal, nPay, nFBal As Double

Dim datatable As New ADODB.Recordset

cStr1 = " SELECT FILE1_10.SUPLER,FILE4_10.DESCA, Sum(FILE6_20.TOTAL) AS TTOTAL, Sum(FILE6_20.QUANT) AS  TQUANT , Sum(Val(FILE6_20.QUANT & '') * Val(file6_20.cost & '') ) AS  Tcost " & _
        " FROM ((FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO)  LEFT JOIN FILE4_10 ON FILE1_10.SUPLER = FILE4_10.CODE "
If IsDate(xdate1.Text) Then cStr1 = cStr1 & turnFound2(cStr1) & " FILE6_20H.DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cStr1 = cStr1 & turnFound2(cStr1) & "  FILE6_20H.DATE <= " & DateSq(XDATE2.Text)
cStr1 = cStr1 & " GROUP BY FILE1_10.SUPLER,FILE4_10.DESCA "

datatable.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText

With grid1
    I = 1
    grid1.Rows = 1
    Do Until datatable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = datatable!SUPLER & ""
        .TextMatrix(.Rows - 1, 1) = datatable!Desca & ""
        .TextMatrix(.Rows - 1, 2) = Format(Val(datatable!TQUANT & ""), "#0")
        .TextMatrix(.Rows - 1, 3) = Format(Val(datatable!TTOTAL & ""), "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(Val(datatable!Tcost & ""), "#0.00")
        If Val(.TextMatrix(.Rows - 1, 4)) > 0 Then .TextMatrix(.Rows - 1, 5) = Format((Val(.TextMatrix(.Rows - 1, 3)) - Val(.TextMatrix(.Rows - 1, 4))) / Val(.TextMatrix(.Rows - 1, 4)) * 100, "#0.00") & "%"
        datatable.MoveNext
    Loop
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 3, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 4, "#0", , RGB(255, 0, 0), True, " "
End With
datatable.Close
Set datatable = Nothing
Me.MousePointer = 0
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
'   mydb.Execute " UPDATE FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM SET FILE6_20.cost = [file1_10].[cost] "
End Sub
Private Sub Form_Load()
With grid1
    .ExplorerBar = flexExSortShow
    .Editable = flexEDNone
    .Cols = 6
    .Rows = 1
    
    .FixedRows = 1

    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next I
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "≈”„"
    .TextMatrix(0, 2) = "„»Ì⁄« \⁄œœ ﬁÿ⁄ "
    .TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄«  "
    .TextMatrix(0, 4) = " ﬂ·›… „»Ì⁄«  "
    .TextMatrix(0, 5) = "‰”»… «·—»Õ"
End With
xdate1.Text = "1-1-2006"
XDATE2.Text = Format(Date, "YYYY-MM-DD")
End Sub

Private Sub Grid1_DblClick()
    ViewITEMSAL.Show 1
End Sub
