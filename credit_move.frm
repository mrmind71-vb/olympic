VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form creditMovefrm 
   Caption         =   "Õ—ﬂ… „œÌ‰Ê‰"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
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
   ScaleHeight     =   10455
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   5595
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3330
         Picture         =   "credit_move.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "credit_move.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4425
         Picture         =   "credit_move.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2235
         Picture         =   "credit_move.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1140
         Picture         =   "credit_move.frx":9573
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   45
      Width           =   5685
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
         Left            =   2430
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   2310
      End
      Begin VB.TextBox xDate2 
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
         Left            =   90
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo xCode 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„œÌ‰"
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
         Height          =   270
         Left            =   4845
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
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
         Height          =   270
         Left            =   4875
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   2355
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
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
         Height          =   270
         Left            =   1635
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1920
      Width           =   405
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Height          =   8250
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   15000
      _cx             =   26458
      _cy             =   14552
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
      Cols            =   8
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
      WordWrap        =   -1  'True
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
End
Attribute VB_Name = "creditMovefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim DocTable As ADODB.Recordset
Dim formMode
Sub fillgrd()
Dim nPrevious As Double
grid1.Rows = 1
xBal.Caption = ""
If DocTable.State = adStateOpen Then DocTable.Close
If IsDate(xdate1.Text) Then
    Dim loctable As New ADODB.Recordset
    loctable.Open "select sum(value_M - value_P) as Balance from credit_move where " & _
                  " CODE = " & MyParn(xCode.BoundText) & _
                  " and [date] < " & DateSq(xdate1.Text), con, adOpenStatic, adLockReadOnly
    If Not loctable.EOF Then nPrevious = Val(loctable!BALANCE & "")
    If nPrevious <> 0 Then
        grid1.AddItem ""
        grid1.TextMatrix(grid1.Rows - 1, 0) = Format(DateAdd("d", -1, xdate1.Text), "YYYY-MM-DD")
        grid1.TextMatrix(grid1.Rows - 1, 2) = "—’Ìœ ”«»ﬁ"
        grid1.TextMatrix(grid1.Rows - 1, 7) = Format(loctable!BALANCE, "#0.00")
    End If
    loctable.Close
    Set loctable = Nothing
End If

cString = "Select  credit_move.*,FILE0_50.DESCA AS BOX_DESCA from credit_move LEFT JOIN FILE0_50 ON credit_move.BOX = FILE0_50.CODE  Where credit_move.CODE = " & MyParn(xCode.BoundText)
If IsDate(xdate1.Text) Then cString = cString & " and date >= " & DateSq(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " and date <= " & DateSq(xDate2.Text)

cString = cString & " Order by [Date],VALUE_P,VALUE_M"
DocTable.Open cString, con, adOpenStatic, , adCmdText

If DocTable.EOF And DocTable.BOF Then Exit Sub
With grid1
Do
   grid1.AddItem ""
   .TextMatrix(.Rows - 1, 0) = Format(DocTable![Date], "YYYY-MM-DD")
   .TextMatrix(.Rows - 1, 1) = DocTable!doc_no & ""
   .TextMatrix(.Rows - 1, 2) = DocTable!TypeDesca & ""
   .TextMatrix(.Rows - 1, 3) = DocTable!Desca & ""
   .TextMatrix(.Rows - 1, 4) = DocTable!Box_Desca & ""
   .TextMatrix(.Rows - 1, 5) = Myvalue(DocTable!Value_M, "FIXED")
   .TextMatrix(.Rows - 1, 6) = Myvalue(DocTable!Value_P, "FIXED")
   .TextMatrix(.Rows - 1, 7) = Format(nPrevious + Val(DocTable!Value_M & "") - Val(DocTable!Value_P & ""), "FIXED")
   nPrevious = Round(nPrevious + Val(DocTable!Value_M & "") - Val(DocTable!Value_P & ""), 2)
   DocTable.MoveNext
   i = i + 1
Loop Until DocTable.EOF
xBal.Caption = nPrevious
End With
End Sub
Sub myProc()
'ActiveControl.Text = GrdText(Search3.Grid1, 0)
Unload Search
End Sub
Function MYVALID()
If xCode.Text = "" Then Exit Function
MYVALID = True
End Function

Private Sub cmdExel_Click()
ToFileExel grid1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
fillgrd
End Sub
Private Sub cmdNew_Click()
grid1.Rows = 1
cmdNew.Enabled = False
cmdGo.Enabled = True
xCode.Enabled = True
xCodeName.Caption = ""
xCode.SetFocus
End Sub
Private Sub cmdPrint_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
Dim aHeader As Variant
cHeader1 = " ›’Ì·Ì Õ—ﬂ… «·„œÌ‰ " & xCode.Text & "  Œ·«· › —…"
If IsDate(xdate1.Text) Or IsDate(xDate2.Text) Then aHeader = AddFlag(aHeader, BetweenString(Format(xdate1.Text, "YYYY-MM-DD"), xDate2.Text))
If Not IsEmpty(aHeader) Then
    cHeader2 = retHeader(aHeader, 0, 1)
End If
PrintGrdNew.doprint grid1, 0.8, -3, cHeader1, cHeader2, cHeader3, , False, False, 9
PrintGrdNew.Show 1
End Sub
Private Sub Form_Load()
openCon con
Set DocTable = New ADODB.Recordset
data1.ConnectionString = strCon
data1.RecordSource = "FILE8_101"

Set xCode.RowSource = data1
xCode.ListField = "Desca"
xCode.BoundColumn = "code"

grid1.FormatString = " «—ÌŒ |" & "—ﬁ„ «·„” ‰œ|" & "‰Ê⁄ |" & "»Ì«‰ |" & "«·Œ“‰…|" & "„œÌÊ‰Ì…|" & "”œ«œ „œÌÊ‰Ì…|" & "—’Ìœ"
grid1.ColWidth(0) = 1500
grid1.ColWidth(1) = 1200
grid1.ColWidth(2) = 2500
grid1.ColWidth(3) = 4000
grid1.ColWidth(4) = 1300
grid1.ColWidth(5) = 1300
grid1.ColWidth(6) = 1300
grid1.ColWidth(7) = 1300
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not xCode.MatchedWithList Then xCode.BoundText = ""
End Sub
Private Sub LastOne_GotFocus()
myGotFocus LastOne
End Sub
Private Sub LastOne_LostFocus()
myLostFocus LastOne
End Sub
