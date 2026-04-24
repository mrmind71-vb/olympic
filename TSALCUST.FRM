VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form TSalCust 
   BackColor       =   &H00FFFFFF&
   Caption         =   "√” «– «·⁄„·«¡ 000   ≈Ã„«·Ï „»Ì⁄«  ··⁄„·«¡"
   ClientHeight    =   7260
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9390
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
   ScaleHeight     =   7260
   ScaleWidth      =   9390
   Begin Crystal.CrystalReport Report1 
      Left            =   720
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "⁄—÷ «Ã„«·Ì „»Ì⁄«  «·«’‰«› ·⁄„Ì·"
      Height          =   315
      Index           =   1
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2940
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "⁄—÷ «Ã„«·Ì ›Ê« Ì— «·⁄„Ì·"
      Height          =   315
      Index           =   0
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   825
      Width           =   2940
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00C0FFFF&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   3300
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄—÷"
      Height          =   390
      Left            =   4875
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton Cmd_Print 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÿ»«⁄…"
      Height          =   390
      Left            =   1725
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   1515
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
      Height          =   315
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   1440
   End
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
      Height          =   315
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1440
   End
   Begin VSFlex7LCtl.VSFlexGrid InvGrid 
      Height          =   4875
      Left            =   360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2160
      Width           =   8835
      _cx             =   15584
      _cy             =   8599
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   89
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TSalCust.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   5115
      Left            =   75
      Top             =   2100
      Width           =   9240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   1365
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   9090
   End
   Begin VB.Label Label2 
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
      Left            =   8175
      TabIndex        =   4
      Top             =   1275
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
      Left            =   8175
      TabIndex        =   3
      Top             =   825
      Width           =   675
   End
End
Attribute VB_Name = "TSalCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim invtable As Recordset, nOption As Byte
Dim TempTable As Recordset
Dim cString As String
Dim cString2 As String
Dim cFilter1 As String
Dim cFilter As String
Dim nInd As Byte
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With invtable
If .RecordCount > 0 Then
.MoveFirst
Do While Not .EOF
    TempTable.AddNew
    TempTable.str1 = .code
    TempTable.val1 = .SUMSAL
    TempTable.str2 = .F_DescA
    If lRepCust Then
        TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï „»Ì⁄«  «·⁄„·«¡ „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.STR7 = "ﬁÌ„… «·„»Ì⁄« "
        TempTable.VAL2 = .SUMSAL - .SUMCOST
    Else
        TempTable.STR6 = "»Ì«‰ »≈Ã„«·Ï „‘ —Ì«  „‰ «·„Ê—œÌ‰  „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xDate2.Text
        TempTable.STR7 = "ﬁÌ„… «·„‘ —Ì« "
    End If
'    TempTable.str19 = firsttitle
'    TempTable.str20 = Secondtitle
    TempTable.Update
    .MoveNext
Loop
End If
End With
REPORT1.ReportFileName = PublicPath & "\Reports\TSalCust.rpt"
REPORT1.DataFiles(0) = App.Path & "\temp.mdb"
REPORT1.Action = 1
Exit Sub
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalCust = Nothing
End Sub
Private Sub CmdOk_Click()
Me.MousePointer = 11
cFilter = ""
If lRepCust Then
    cString = "SELECT FILE3_10.CODE, First(FILE3_10.DESCA) AS F_DESCA, Sum(FILE6_20.TOTAL) AS SumSAL,  " & _
                " Sum([COST]*[QUANT]) AS SumCost, FILE3_10.group, FILE1_70.DESCA as GrDesc " & _
                " FROM (FILE3_10 RIGHT JOIN FILE6_20 ON FILE3_10.CODE = FILE6_20.CODE) LEFT JOIN FILE1_70 ON  " & _
                " FILE3_10.group = FILE1_70.CODE " & _
                " WHERE FILE1_70.FLAG =3 "
                If xDate1.Text <> "" Then
                    cString = cString & " AND Date >= DateValue('" & xDate1.Text & "')"
                End If
                If xDate2.Text <> "" Then
                    cString = cString & " AND Date <= DateValue('" & xDate2.Text & "')"
                End If
                cString = cString & " GROUP BY FILE3_10.CODE, FILE3_10.group, FILE1_70.DESCA "
                cString = cString & " ORDER BY FILE3_10.GROUP , FILE3_10.CODE"
Else
    cString = "SELECT FILE4_10.CODE, First(FILE4_10.DESCA) AS F_DESCA, Sum(FILE7_20.TOTAL) AS SumSAL,  " & _
                " Sum([COST]*[QUANT]) AS SumCost, FILE4_10.group, FILE1_70.DESCA as GrDesc " & _
                " FROM (FILE4_10 RIGHT JOIN FILE7_20 ON FILE4_10.CODE = FILE7_20.CODE) LEFT JOIN FILE1_70 ON  " & _
                " FILE4_10.group = FILE1_70.CODE " & _
                " WHERE FILE1_70.FLAG =4 "
                If xDate1.Text <> "" Then
                    cString = cString & " AND Date >= DateValue('" & xDate1.Text & "')"
                End If
                If xDate2.Text <> "" Then
                    cString = cString & " AND Date <= DateValue('" & xDate2.Text & "')"
                End If
                cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.group, FILE1_70.DESCA "
                cString = cString & " ORDER BY FILE4_10.GROUP , FILE4_10.CODE"
End If
Set invtable = mydb.CreateSnapshot(cString)
If invtable.RecordCount > 0 Then
    Fillgrd
Else
    InvGrid.Rows = 1
    MsgBox ("·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«")
End If
Me.MousePointer = 0
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'lCust = (ption1(0).Value)
If lRepCust Then
    Me.Caption = "√” «– «·⁄„·«¡ 000   ≈Ã„«·Ï „»Ì⁄«  ··⁄„·«¡"
Else
    Me.Caption = "√” «– «·„Ê—œÌ‰ 000   ≈Ã„«·Ï „‘ —Ì«  „‰ «·„Ê—œÌ‰"
End If
With InvGrid
.Cols = 6
.Rows = 1
.TextMatrix(0, 1) = "ﬂÊœ"
.TextMatrix(0, 2) = "„Ã„Ê⁄…"
.TextMatrix(0, 3) = "ﬂÊœ"
.TextMatrix(0, 4) = "⁄„Ì·"
.TextMatrix(0, 5) = "Ã. ﬁÌ„… „»Ì⁄« "
.ColWidth(0) = 400
.ColWidth(1) = 0
.ColWidth(2) = 1500
.ColWidth(3) = 0
.ColWidth(4) = 3000
.ColWidth(5) = 1500
For i = 0 To InvGrid.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.MergeCells = 2
End With

Option1(0).Value = True
nInd = 1
xDate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub InvGrid_Click()
End Sub
Private Sub xInd_Click(Index As Integer)
    If xInd(Index).Value = True Then nInd = Index
End Sub
Sub Fillgrd()
With InvGrid
.FixedRows = 1
.FixedCols = 1
.ExplorerBar = flexExSortShow
InvGrid.Rows = 1
invtable.MoveFirst
.MergeCells = 2
.SubtotalPosition = flexSTBelow
i = 1
Do
   .AddItem ""
   .TextMatrix(i, 1) = invtable.Group
   .TextMatrix(i, 2) = invtable.grdesc
   .TextMatrix(i, 3) = invtable.code
   .TextMatrix(i, 4) = invtable.F_DescA
   .TextMatrix(i, 5) = Format(TurnValue(invtable.SUMSAL, Null, 0) - TurnValue(invtable.SUMCOST, Null, 0), "##0.00")
   .MergeCol(1) = True
   .MergeCol(2) = True
   .MergeCol(3) = True
   .MergeCol(4) = True
   .MergeCol(5) = True
   .MergeRow(i) = True
   invtable.MoveNext
   i = i + 1
Loop Until invtable.EOF
.Sort = 1
.Subtotal flexSTClear
.Subtotal flexSTSum, -1, 5, , , RGB(255, 0, 0), True
.Subtotal flexSTSum, 1, 5, , , RGB(155, 100, 0), True
.OutlineBar = 1
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
.Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbBlue

End With
End Sub

