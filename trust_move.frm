VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form trustMovefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÊÌ«  ”«∆ﬁÌ‰"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   6675
      Begin VB.TextBox XCODE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4320
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4320
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XDATE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2745
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂÊœ"
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
         Height          =   270
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ"
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
         Height          =   270
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   270
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1215
         Picture         =   "trust_move.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "trust_move.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "trust_move.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "trust_move.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4410
      Top             =   585
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
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
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -645
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1965
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7575
      Left            =   90
      TabIndex        =   12
      Top             =   1035
      Width           =   14775
      _cx             =   26061
      _cy             =   13361
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
      Cols            =   7
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   8700
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "trustMovefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, oSearch As New Search3
Dim ClientTable As New ADODB.Recordset
Sub fillgrd()
Dim loctable As New ADODB.Recordset
cString = "SELECT TRUST_MOVE.*  " & _
          " From TRUST_MOVE"

cString = cString & turn(cString) & " TRUST_MOVE.BOX = " & MyParn(XCODE.Text)
If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " TRUST_MOVE.date >= " & DateSq(xdate1.Text)
End If

If IsDate(XDATE2.Text) Then
    cString = cString & turn(cString) & " TRUST_MOVE.date <= " & DateSq(XDATE2.Text)
End If
cString = cString & " order by Date,TRUST_MOVE.DOC_NO,TRUST_MOVE.PLUS"

With grid1
    .Rows = 1
    If IsDate(xdate1.Text) Then
       cString2 = "Select sum([PLUS] - MINUS) as Balance from TRUST_MOVE where TRUST_MOVE.BOX = " & MyParn(XCODE.Text) & _
                  " and TRUST_MOVE.DATE < " & DateSq(xdate1.Text)
       nPrevious = Val(GetField(cString2) & "")
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Val(loctable!PLUS & "") - Val(loctable!MINUS & "")
         If Not IsNull(loctable!Notes) Then
            .TextMatrix(.Rows - 1, 0) = loctable!Notes & ""
         Else
            .TextMatrix(.Rows - 1, 0) = loctable!type_Desca & ""
         End If
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!DOC_NO & ""
        .TextMatrix(.Rows - 1, 3) = Myvalue(loctable!PLUS, "FIXED")
        .TextMatrix(.Rows - 1, 4) = Myvalue(loctable!MINUS, "FIXED")
        .TextMatrix(.Rows - 1, 5) = Myvalue(nPrevious, "FIXED")
        .TextMatrix(.Rows - 1, 6) = loctable!Type & ""
        loctable.MoveNext
    Loop
    
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
    If grid1.Rows > 1 Then
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 5) = Format(Round(nPrevious, 2), "#0.00")
    End If
End With
StatusBar1.Panels(1).Text = "—’Ìœ «·”«∆ﬁ :  " & Format(nPrevious, "#0.00")
End Sub
Sub myProc()
XCODE.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xDesca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Function MYVALID() As Boolean
If Trim(XCODE.Text) = "" Then
    MsgBox "ﬂÊœ «·”«∆ﬁ €Ì— „”Ã·"
    Exit Function
End If
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(XDATE2.Text)) And Trim(XDATE2.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub

Private Sub cmdExel_Click()
ToFileExel grid1
End Sub

Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
fillgrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String
Dim aHeader As Variant
cHeader1 = " ”ÊÌ«  ”«∆ﬁ Œ·«· › —…"
If IsDate(xdate1.Text) Or IsDate(XDATE2.Text) Then aHeader = AddFlag(aHeader, BetweenString(xdate1.Text, XDATE2.Text))
If Trim(XCODE.Text) <> "" Then aHeader = AddFlag(aHeader, "··”«∆ﬁ : " & xDesca.Caption)
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", grid1.Rows - 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 2)
PrintGrdNew.doprint grid1, 0.9, -1, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 9, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
openCon con
With grid1
grid1.Cols = 7
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "œ«∆‰"
.TextMatrix(0, 4) = "„œÌ‰"
.TextMatrix(0, 5) = "—’Ìœ"

grid1.ColWidth(0) = 3000
grid1.ColWidth(1) = 1500
grid1.ColWidth(2) = 1500
grid1.ColWidth(3) = 1500
grid1.ColWidth(4) = 1500
grid1.ColWidth(5) = 1500
grid1.ColHidden(6) = True

End With
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub
Private Sub Grid1_DblClick()
Select Case grid1.TextMatrix(grid1.Row, 6)
    Case "1"
        trustfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 2)
        trustfrm.Show
    Case "2"
        trust_cashfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 2)
        trust_cashfrm.Show
End Select
End Sub

Private Sub xCode_Validate(Cancel As Boolean)
xDesca.Caption = ""
If Trim(XCODE.Text) = "" Then Exit Sub
XCODE.Text = RetZero(XCODE.Text, 6)
Dim aRet As Variant
aRet = GetField("SELECT DESCA FROM DRIVER WHERE DRIVER = 1 AND CODE = " & MyParn(XCODE.Text))
If IsEmpty(aRet) Then
    MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xDesca.Caption = aRet & ""
End If
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xCode_Change()
grid1.Rows = 1
cmdGo.Enabled = Trim(XCODE.Text) <> ""
End Sub
Private Sub XCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then DriverLookupAll Me, oSearch, "DRIVER = 1"
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate XDATE2
End Sub
