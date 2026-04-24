VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form supMovefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·„Ê—œÌ‰"
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
         Left            =   3600
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
         Left            =   3600
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
         Left            =   2025
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
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   3300
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·„Ê—œ"
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
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   825
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
         Left            =   5280
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
         Picture         =   "supMove.frx":0000
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
         Picture         =   "supMove.frx":27EB
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
         Picture         =   "supMove.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2415
         Picture         =   "supMove.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1980
      Top             =   0
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
Attribute VB_Name = "supMovefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ClientTable As New ADODB.Recordset
Sub fillgrd()
Dim loctable As New ADODB.Recordset
cString = "select FILE4_11.*  " & _
          " From FILE4_11 "

cString = cString & turnFound(cString) & " FILE4_11.code = " & MyParn(xCode.Text)

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & " FILE4_11.date >= " & DateSq(xdate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & " FILE4_11.date <= " & DateSq(xDate2.Text)
End If

cString = cString & " order by Date,file4_11.doc_id,file4_11.sal"

With grid1
    .Rows = 1
    If IsDate(xdate1.Text) Then
       cString2 = "Select sum([SAL] - PAY) as Balance from FILE4_11 where FILE4_11.CODE = " & MyParn(xCode.Text) & _
                  " and FILE4_11.date < " & DateSq(xdate1.Text)
       nPrevious = Val(GetDesca(cString2))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Val(loctable!sal & "") - Val(loctable!Pay & "")
        .TextMatrix(.Rows - 1, 0) = loctable!type_Desca & ""
        If Not IsNull(loctable!Desca) Then
            .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 1, 0) & turn(.TextMatrix(.Rows - 1, 0), " ") & "[" & loctable!Desca & "]"
        End If
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!doc_ID & ""
        .TextMatrix(.Rows - 1, 3) = Format(TurnValue(Val(loctable!sal & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(TurnValue(Val(loctable!Pay & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 5) = Format(nPrevious, "#0.00")
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
StatusBar1.Panels(1).Text = "—’Ìœ «·„Ê—œ :  " & Format(nPrevious, "#0.00")
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— „”Ã·"
    Exit Function
End If
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then
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
cHeader1 = "Õ—ﬂ… „Ê—œ Œ·«· › —…"
If IsDate(xdate1.Text) Or IsDate(xDate2.Text) Then aHeader = AddFlag(aHeader, BetweenString(xdate1.Text, xDate2.Text))
If Trim(xCode.Text) <> "" Then aHeader = AddFlag(aHeader, "··„Ê—œ : " & xDesca.Caption)
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", grid1.Rows - 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", 2)
PrintGrdNew.doprint grid1, 0.9, -1, cHeader1, retHeader(aHeader, 0, 1), retHeader(aHeader, 1, 2), , False, False, 9, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
openCon con
ClientTable.Open "FILE4_10", con, adOpenStatic, adLockReadOnly, adCmdTable

With grid1
grid1.Cols = 7
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "œ«∆‰"
.TextMatrix(0, 4) = "„œÌ‰"
.TextMatrix(0, 5) = "—’Ìœ"

grid1.ColWidth(0) = 5000
grid1.ColWidth(1) = 1400
grid1.ColWidth(2) = 1300
grid1.ColWidth(3) = 1300
grid1.ColWidth(4) = 1300
grid1.ColWidth(5) = 1300
grid1.ColHidden(6) = True

End With
For I = 0 To grid1.Cols - 1
    grid1.ColAlignment(I) = flexAlignRightCenter
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub
Private Sub Grid1_DblClick()
'    Select Case grid1.TextMatrix(grid1.row, 6)
'        Case "4", "5"
'            Load Purchasefrm
'            Purchasefrm.myPublic = IIf(grid1.TextMatrix(grid1.row, 6) = "4", 0, 1)
'            Purchasefrm.myproc2 grid1.TextMatrix(grid1.row, 2)
'            Purchasefrm.Frame9.Visible = False
'            Purchasefrm.Frame1.Visible = False
'            Purchasefrm.Show 1
'        Case "A", "C"
'            publicFlag = 2
'            Load chq
'            chq.MoveLOadChq grid1.TextMatrix(grid1.row, 2)
'            chq.Show 1
'    End Select
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub

Private Sub xCode_Change()
grid1.Rows = 1
cmdGo.Enabled = Trim(xCode.Text) <> ""
End Sub

Private Sub XCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xDesca.Caption = ""
If Trim(xCode.Text) = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
ClientTable.Find "code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not ClientTable.EOF Then xDesca.Caption = ClientTable!Desca & ""
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE4_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

Private Sub doprint()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim cStrW As String

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If Trim(xCode.Text) <> "" Then
    aHeader(0) = "[" & "··„Ê—œ : " & xDesca.Caption & "]"
End If
If IsDate(xdate1.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
    cStrW = cStrW & " AND DATE >= " & DateSq(xdate1.Text)
End If
If IsDate(xDate2.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
    cStrW = cStrW & " AND DATE <= " & DateSq(xDate2.Text)
End If

With grid1
For I = 1 To .Rows - 2
    temptable.AddNew
    temptable!Date1 = TurnValue(DateFix(.TextMatrix(I, 1)))
    temptable!str1 = TurnValue(.TextMatrix(I, 2))
    temptable!str2 = TurnValue(.TextMatrix(I, 0))
    temptable!val1 = Val(.TextMatrix(I, 3))
    temptable!val2 = Val(.TextMatrix(I, 4))
    temptable!Val3 = Val(.TextMatrix(I, 5))
    temptable!Val6 = I
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!STR20 = Firsttitle
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\sup_move.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub LastOne_GotFocus()
myGotFocus LastOne
End Sub
Private Sub LastOne_LostFocus()
myLostFocus LastOne
End Sub

