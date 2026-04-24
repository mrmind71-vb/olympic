VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form cardTransfrm 
   Caption         =   "Œÿ«»«  «‰–«—"
   ClientHeight    =   10110
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   16785
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   16785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2745
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   6225
      Begin VB.CommandButton cmdPrintLetter 
         Appearance      =   0  'Flat
         Caption         =   "«—”«· «·»Ì«‰«  ··»Ê«»…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3195
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1905
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   5130
         Picture         =   "card_trans.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2145
         Picture         =   "card_trans.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "card_trans.frx":491C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1095
         Picture         =   "card_trans.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1050
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9735
      Width           =   16785
      _ExtentX        =   29607
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   7710
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   225
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „Ê”„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   450
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8295
      Left            =   90
      TabIndex        =   1
      Top             =   765
      Width           =   16620
      _cx             =   29316
      _cy             =   14631
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
      Cols            =   9
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
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "cardTransfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Dim aHeader()

Private Sub cmdExel_Click()
Dim sHeader As String, nMargin As Integer

sHeader = Me.Caption
nMargin = 40
If retHeader(aHeader, 1, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 1, 3)
    nMargin = nMargin + 15
End If



Dim aSplit As Variant
aSplit = AddFlag(aSplit, "title_col", "A:B")
aSplit = AddFlag(aSplit, "title_row", "1:1")
aSplit = AddFlag(aSplit, "center_header", sHeader)
ToFileExel grid1, , , , , , , , aSplit, , , , nMargin

'ToFileExel grid1
End Sub
Private Sub CmdPrint_Click()
PrintGrdNew.doprint grid1, 0.85, 0, Me.Caption, retHeader(aHeader, 1, 2), retHeader(aHeader, 3, 2), , False, False, 10
PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
myload
End Sub
Private Sub cmdPrintLetter_Click()
doprint
End Sub

Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub
Private Sub Form_Load()
Me.Top = 1000
Me.Left = 1000
openCon con

Set grid1.DataSource = DATA10
Fixgrd
LoadText Me
If ValidNum(cmdYear(1).Tag) Then
    cmdYear(1).Caption = GetField("select desca from years_codes where code = " & cmdYear(1).Tag, con) & ""
End If
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(4)
With grid1
cString = "SELECT FILE1_10.CODE,FILE1_10.DESCA,dbo.f_last_year_date(file1_10.code) as date_last,dbo.f_save(FILE1_10.CODE) AS IS_SAVE" & _
          " FROM FILE1_10"
cWhere = "FILE1_10.[dropped] = 0"
cWhere = cWhere & turn(cWhere, " and ") & "dbo.f_last_year_code(file1_10.code) >= " & cmdYear(1).Tag

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If cmdYear(1).Tag <> "" Then
    aHeader(3) = "··«⁄÷«¡ «·„”œœÌ‰ „Ê”„ : " & cmdYear(1).Caption
End If

cString = cString & " ORDER BY FILE1_10.CODE"
Set DATA10.Recordset = myRecordSet(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim i As Long
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 2) = "«”„ «·⁄÷Ê"
    .TextMatrix(0, 3) = "«Œ— ”œ«œ"
    .TextMatrix(0, 4) = "Õ«›Ÿ ⁄÷ÊÌ…"
        
    .ColWidth(0) = 800
    .ColWidth(1) = 1000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColDataType(3) = flexDTDate
    .ColDataType(4) = flexDTBoolean
    
    For i = 1 To grid1.rows - 1
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 6) = myFormat_p(.TextMatrix(i, 6))
    Next
    
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    SBar1.Panels(1).Text = IIf(grid1.rows > 2, "⁄œœ «·”Ã·«  : " & grid1.rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me, , Array(xDate.Name, cmdYear(1).Name)
closeCon con
Set cardTransfrm = Nothing
End Sub
Private Sub grid1_DblClick()
With grid1
End With
End Sub
Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidInt(xCode.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode.Text)
If Not IsEmpty(aRet) Then xCodeDesca.Caption = retFlag(aRet, "DESCA") & ""
End Sub

Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
'cmdPrint.Enabled = grid1.rows > 1
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xbox_GotFocus()
myGotFocus xbox
End Sub
Private Sub xbox_LostFocus()
myLostFocus xbox
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(1).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Function MYVALID() As Boolean
If cmdYear(1).Tag = "" Then
    MsgBox "·« ÌÊÃœ „Ê”„"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    MsgBox IIf(Trim(xDate.Text) = "", " «—ÌŒ €Ì— „”Ã·", " «—ÌŒ €Ì— ’ÕÌÕ")
    Exit Function
End If
MYVALID = True
End Function
Private Function getYears(pCode As String, pDate As String) As String
Dim loctable As New ADODB.Recordset
cString = "select [year] from years_codes where code <=  dbo.f_ret_year(" & DateSq(pDate) & ") and code > " & cmdYear(1).Tag & "  group by [year] order by [year] DESC"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    getYears = getYears & turn(getYears, "°") & loctable!Year
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function
Private Function doprint()
Dim temptable As New ADODB.Recordset, cOr As String


contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With grid1
For i = 1 To grid1.rows - 1
    If Trim(.TextMatrix(i, 4)) <> "" Then
        temptable.AddNew
        temptable!val1 = TurnValue(.TextMatrix(i, 1))
        temptable!str1 = TurnValue(ArbString(.TextMatrix(i, 1)))
        temptable!str2 = TurnValue(.TextMatrix(i, 2))
        temptable!str16 = TurnValue(ArbString(.TextMatrix(i, 4)))
        If Trim(.TextMatrix(i, 3)) <> "" Then
            temptable!str5 = TurnValue("/" & .TextMatrix(i, 3))
        Else
            temptable!str5 = "«·”«œ…/"
        End If
        temptable!str4 = TurnValue(ArbString(aHeader(0)))
        temptable.Update
    End If
Next
End With

temptable.Requery
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    REPORT1.ReportFileName = sPath_App & "\REPORTS\warrent.rpt"
    REPORT1.DataFiles(0) = tempFile
    REPORT1.Action = 1
End If

Set temptable = Nothing
End Function


Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
