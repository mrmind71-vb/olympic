VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form VsBalItem 
   Caption         =   "≈Ã„«·Ï ﬁÌ„… „»Ì⁄«  √’‰«› «·„Ê—œÌ‰"
   ClientHeight    =   9750
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   4335
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
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
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
         Height          =   420
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton CmdOk1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   420
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1365
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grid1 
      Height          =   8955
      Left            =   -90
      TabIndex        =   8
      Top             =   720
      Width           =   15240
      _cx             =   26882
      _cy             =   15796
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
      TabIndex        =   3
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x2 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   8250
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label x1 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   8250
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label x4 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   8640
      Width           =   240
   End
End
Attribute VB_Name = "VsBalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Print_Click()
    Dim cHead As String
    Dim cHead2 As String
    cHead = firsttitle
    cHead2 = "»Ì«‰ »⁄œœ Ê ﬁÌ„… —’Ìœ √’‰«› „Ê—œÌ‰ "
    Load PrintGrd
    PrintGrd.doprint Grid1, 0.9, -2, cHead, cHead2, , False, , 8
    PrintGrd.Show 1
End Sub
Private Sub CMDEXIT_Click()
Unload Me
Set VsBalItem = Nothing
End Sub
Private Sub CmdOk1_Click()
Dim nBal1 As Double, nBal2 As Double, nBal3 As Double
Dim nTot As Double
Dim nSal, nPay, nFBal As Double
Dim datatable As New ADODB.Recordset

cStr1 = " SELECT FILE1_10.SUPLER, Sum(FILE1_11.OUT) AS TOUT, Sum(FILE1_11.[IN]) AS  TIN , Sum( VAL(FILE1_11.OUT & '') * VAL(FILE1_10.cost & '') ) AS VOUT , Sum(VAL(FILE1_11.[IN] & '') * VAL(FILE1_10.cost & '') ) AS  VIN , Sum(VAL(FILE1_11.OUT & '') * VAL(FILE1_10.PRICE & '')) AS VOUT2 , Sum(VAL(FILE1_11.[IN] & '') * VAL(FILE1_10.PRICE & '')) AS  VIN2,FILE4_10.DESCA " & _
        " FROM (FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) INNER JOIN FILE4_10 ON  FILE1_10.SUPLER =  FILE4_10.CODE "

cStr1 = cStr1 & " GROUP BY FILE1_10.SUPLER,FILE4_10.DESCA "
datatable.Open cStr1, CON, adOpenStatic, adLockReadOnly, adCmdText


With Grid1
    i = 1
    Grid1.Rows = 1
    Do Until datatable.EOF
        .AddItem ""
        .TextMatrix(i, 0) = datatable!SUPLER & ""
        .TextMatrix(i, 1) = datatable!desca & ""
        
        .TextMatrix(i, 2) = Format(Val(datatable!TIN & "") - Val(datatable!TOUT & ""), "#0")
        .TextMatrix(i, 3) = Format(Val(datatable!VIN & "") - Val(datatable!VOUT & ""), "#0")
        .TextMatrix(i, 4) = Format(Val(datatable!vIN2 & "") - Val(datatable!vOUT2 & ""), "#0")
        If Val(.TextMatrix(i, 3)) > 0 Then .TextMatrix(i, 5) = Format((Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 3))) / Val(.TextMatrix(i, 3)) * 100, "#0.00")
        datatable.MoveNext
        i = i + 1
    Loop
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 3, "#0", , RGB(255, 0, 0), True, " "
    .Subtotal flexSTSum, -1, 4, "#0", , RGB(255, 0, 0), True, " "
End With
Me.MousePointer = 0
datatable.Close
Set datatable = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()

With Grid1
    .ExplorerBar = flexExSortShow
    .Editable = flexEDNone
    .Cols = 6
    .Rows = 1
    
    .FixedRows = 1

    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next i
    .ColWidth(0) = 1000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "≈”„"
    .TextMatrix(0, 2) = "—’Ìœ ⁄œœ ﬁÿ⁄ "
    .TextMatrix(0, 3) = "ﬁÌ„… —’Ìœ  ﬂ·›…"
    .TextMatrix(0, 4) = "ﬁÌ„… —’Ìœ »”⁄— «·»Ì⁄ "
    .TextMatrix(0, 5) = "‰”»… «·„” Â·ﬂ "
End With
End Sub
Private Sub grid1_DblClick()
    ViewITEMBAL.Show 1
End Sub
