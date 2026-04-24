VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chargefrm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„’«—Ì›"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18825
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9600
   ScaleWidth      =   18825
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkPrv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«ŸÂ«—  «—ÌŒ «Œ— «’·«Õ"
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
      Height          =   240
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "charge1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge1.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge1.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge1.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8595
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   11
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "charge1.frx":9A37
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge1.frx":BC07
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   12
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "charge1.frx":DD4F
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge1.frx":FF17
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   13
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "charge1.frx":12066
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge1.frx":14246
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   14
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "charge1.frx":163A1
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge1.frx":1855D
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   13905
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   675
      Width           =   4830
      Begin VB.TextBox xDate 
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
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1635
      End
      Begin VB.TextBox xDoc_No 
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
         Left            =   2340
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   1275
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge1.frx":1A6AC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge1.frx":1CA0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   9255
      Width           =   18825
      _ExtentX        =   33205
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "03:10 „"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5490
      Top             =   270
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   6765
      Left            =   90
      TabIndex        =   2
      Top             =   1800
      Width           =   18645
      _cx             =   32888
      _cy             =   11933
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
Attribute VB_Name = "chargefrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim bActivated As Boolean
Dim clist1 As String, cList2 As String
Dim CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim oSearchDoc As New Search3, oSearchClient As New Search3, oSearchCar As New Search3, oSearchBox As New Search3
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, cdef_Box As String
Dim formMode
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[USERNAME]", addstring(cUserName))
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, cFileHeader)
Else
    con.Execute addUpdate(aInsert, cFileHeader, "doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(nRow = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "Box", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "Charge", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "Desca", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "[VALUE]", Val(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "[CAR]", addvalue(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "[CODE]", addstring(grid1.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "[COUNTER]", addstring(grid1.TextMatrix(i, 12)))
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, cFile)
        Else
            con.Execute addUpdate(aInsert, cFile, "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
        Unload oSearchBox
    ElseIf grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 1)
        Grid1_AfterEdit grid1.Row, grid1.Col
        Unload oSearchClient
        CellPos 13, grid1.Row, grid1.Col
    ElseIf grid1.Col = 5 Then
        grid1.TextMatrix(grid1.Row, 5) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 6) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 4)
        grid1.TextMatrix(grid1.Row, 7) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 1) & " " & oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 2)
        Unload oSearchCar
        Grid1_AfterEdit grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
End If
End Sub

Private Sub chkPrv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fixgrd
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT " & cFileHeader & ".Doc_No, Convert(Varchar," & cFileHeader & ".Date,111),Min(" & DocClient & ".Desca)" & _
          " FROM (" & cFileHeader & " inner join " & cFile & " on " & cFileHeader & ".doc_no = " & cFile & ".Doc_NO) Inner Join " & DocClient & " on " & cFile & ".Charge = " & DocClient & ".Code"
          
Generalarray(1) = cString
Generalarray(2) = " group by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date order by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%" & DocClient & ".Desca%% or " & _
                  " ##" & cFileHeader & ".Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_Activate()
On Error Resume Next
If Not bActivated Then
    If xDoc_No.Tag = DefineMode Then
        xDate.SetFocus
    Else
        grid1.SetFocus
    End If
    Err.Clear
    bActivated = True
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
openCon con
bedit = True
cFile = "File8_50"
cFileHeader = "FILE8_50H"
DocClient = "FILE8_51"

clist1 = StrBox
cList2 = StrList2("Select CODE,DESCA FROM DRIVER ORDER BY DESCA")
Set grid1.DataSource = data1
data1.ConnectionString = strCon

cdef_Box = myDef("FILE0_50", "CODE")

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
GRDTABLE.Close
Set CardTable = Nothing
Set GRDTABLE = Nothing
closeCon con
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    Calctotals
    Exit Sub
End If
With grid1
If Row = grid1.Rows - 1 Then
    MyAddItem
ElseIf Row = grid1.Rows - 2 And (Col = 3 Or Col = 0) Then
'    MyEditItem grid1, Row, Col
End If
Calctotals
If myreplace(Row) Then
    If xDoc_No.Tag = DefineMode Then
        xDoc_No.Tag = LoadMode
        xDoc_No.Enabled = False
    End If
    If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadgrd
    End If
End If
End With
End Sub

Private Sub grid1_dblClick()
Dim cWhere As String

If grid1.Col = 11 Then
    If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
    If grid1.TextMatrix(grid1.Row, 11) = "" Then Exit Sub
    cWhere = "charge = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
    If IsDate(xDate.Text) Then cWhere = cWhere & turn(cWhere, " and ") & "FILE8_50H.date <=" & DateSq(xDate.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "car = " & grid1.TextMatrix(grid1.Row, 5)
    cWhere = cWhere & turn(cWhere, " and ") & "id <> " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
    showfrm1.cWhere = cWhere
    showfrm1.sDate = xDate.Text
    showfrm1.Show 1
ElseIf grid1.Col = 14 Then
    If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
    If grid1.TextMatrix(grid1.Row, 14) = "" Then Exit Sub
    cWhere = "charge = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
    If IsDate(xDate.Text) Then cWhere = cWhere & turn(cWhere, " and ") & "COUNTER <=" & grid1.TextMatrix(grid1.Row, 12)
    cWhere = cWhere & turn(cWhere, " and ") & "car = " & grid1.TextMatrix(grid1.Row, 5)
    cWhere = cWhere & turn(cWhere, " and ") & "id <> " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
    showfrm3.cWhere = cWhere
    showfrm3.nCounter = Val(grid1.TextMatrix(grid1.Row, 12))
    showfrm3.Show 1
End If
End Sub

Private Sub Grid1_EnterCell()
If grid1.Col = 1 Or grid1.Col = 2 Or grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 5 Or grid1.Col = 12 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 1
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.Text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 1) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
    If Val(.TextMatrix(i, 4)) = 0 Then
        MsgBox "«·ﬁÌ„… €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
Handlecontrols LoadMode
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT " & cFile & ".[BOX], " & cFile & ".Charge," & DocClient & ".DESCA," & cFile & ".desca,[VALUE]," & cFile & ".CAR,CARS.BOARD, CARS.DESCA," & cFile & ".CODE,case when FILE8_51.MONTHES = 0 then NULL else FILE8_51.MONTHES END ,'' AS DATE_PRV,'' AS PERIOD_PRV,COUNTER,'' AS COUNTER_PRV ,'' AS DIFFER_KILO,CASE WHEN FILE8_51.KILOS = 0 THEN NULL ELSE FILE8_51.KILOS END,[ID]  " & _
               " FROM " & cFile & " LEFT JOIN " & DocClient & " ON " & cFile & ".Charge = " & DocClient & ".CODE " & " LEFT JOIN CARS ON " & cFile & ".car = CARS.CODE " & _
               " WHERE " & cFile & ".Doc_no = " & MyParn(xDoc_No.Text)
    data1.RecordSource = cString
    data1.Refresh
    MyAddItem
End With
Calctotals
Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
xDate.Text = Format(Date, "YYYY-MM-DD")
Fixgrd
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = cdef_Box
Handlecontrols DefineMode
Calctotals
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And doc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    BoxLookupAll Me, oSearchBox
ElseIf KeyCode = 112 And grid1.Col = 1 Then
    grdLookup
ElseIf KeyCode = 112 And grid1.Col = 5 Then
    carsLookupAll Me, oSearchCar
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("„‰ «·„” ‰œ ?", vbOKCancel + vbDefaultButton2) = vbOK Then
        con.BeginTrans
'        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        myRemove grid1.Row
        Grid1_EnterCell
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 And Col <> 8 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        grid1.EditText = RetZero(grid1.EditText, 3)
        aRet = GetFields("SELECT CODE,DESCA,DAYS FROM " & DocClient & " WHERE CODE = " & MyParn(grid1.EditText), con)
        If IsEmpty(aRet) Then
           MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ €Ì— ’ÕÌÕ"
           Cancel = True
        Else
            grid1.TextMatrix(Row, 2) = retFlag(aRet, "desca") & ""
            grid1.TextMatrix(Row, 9) = Myvalue(Val(retFlag(aRet, "days") & "") / 30)
        End If
    End If
ElseIf Col = 5 Then
    If Trim(grid1.EditText) = "" Then
        grid1.TextMatrix(grid1.Row, 6) = ""
        grid1.TextMatrix(grid1.Row, 7) = ""
        grid1.TextMatrix(grid1.Row, 8) = ""
    ElseIf (Not ValidInt(grid1.EditText)) Then
        Cancel = True
    Else
        aRet = GetFields("SELECT DESCA,BOARD,MODEL FROM CARS WHERE CODE = " & grid1.EditText)
        If IsEmpty(aRet) Then
            MsgBox "ﬂÊœ «·”Ì«—… €Ì— ’ÕÌÕ"
            Cancel = True
        Else
            grid1.TextMatrix(grid1.Row, 6) = retFlag(aRet, "BOARD")
            grid1.TextMatrix(grid1.Row, 7) = retFlag(aRet, "DESCA")
        End If
    End If
End If
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    boxtable.MoveFirst
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!desca
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From " & DocClient
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchClient.Caption = "≈” ⁄·«„ "
oSearchClient.Show 1
End Sub
Private Function Calctotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 4)), 2)
    If grid1.TextMatrix(i, 13) <> "" Then
        grid1.TextMatrix(i, 14) = Myvalue(Val(grid1.TextMatrix(i, 12)) - Val(grid1.TextMatrix(i, 13)))
        If Val(grid1.TextMatrix(i, 14)) < Val(grid1.TextMatrix(i, 15)) Then grid1.Cell(flexcpBackColor, i, 14, i, 14) = &H8080FF Else grid1.Cell(flexcpBackColor, i, 14, i, 14) = &H80000005
    End If
Next
StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Myvalue(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " Where code = " & MyParn(grid1.TextMatrix(nRow, 1))) & ""
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", con, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbdMouse
    .FormatString = "Œ“‰…|" & "«·„’—Ê›|" & "Ê’› «·„’—Ê›|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "«·”Ì«—…|" & "—ﬁ„ «··ÊÕ…|" & "»Ì«‰ «·”Ì«—…|" & "«·„”∆Ê·|" & "„œ… «·«’·«Õ »«·‘Â—|" & "«Œ—  «—ÌŒ «’·«Õ|" & "«·„œ… »«·‘Â—|" & "ﬁ—«¡… ⁄œ«œ|" & "ﬁ—«¡… ”«»ﬁ…|" & "«·›—ﬁ|" & "›—ﬁ „› —÷|"
    .RowHeight(0) = 600
    .ColWidth(0) = 1800
    .ColWidth(1) = 800
    .ColWidth(2) = 2500
    .ColWidth(3) = 2500
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 800
    .ColWidth(7) = 2000
    .ColWidth(8) = 1500
    .ColWidth(9) = 1000
    .ColWidth(10) = 1400
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    
    '.ColComboList(8) = cList2
    .ColHidden(8) = True
    .ColHidden(0) = True
    .ColHidden(.Cols - 1) = True
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = clist1
    If chkPrv.Value = 1 Then
        For i = 1 To grid1.Rows - 2
            If grid1.TextMatrix(i, 9) <> "" Or grid1.TextMatrix(i, 12) <> "" Then
                If Trim(grid1.TextMatrix(i, 1)) <> "" And Trim(grid1.TextMatrix(i, 5)) <> "" And grid1.TextMatrix(i, .Cols - 1) <> "" Then
                    Dim aRet As Variant
                    cString = "select top 1 file8_50h.date,FILE8_50.COUNTER from file8_50h inner join file8_50 on file8_50h.doc_no = file8_50.doc_no "
                    cString = cString & turn(cString) & "car = " & grid1.TextMatrix(i, 5)
                    cString = cString & turn(cString) & "charge = " & MyParn(grid1.TextMatrix(i, 1))
                    cString = cString & turn(cString) & "file8_50h.date <= " & DateSq(xDate.Text)
                    cString = cString & turn(cString) & "id <> " & grid1.TextMatrix(i, .Cols - 1)
                    cString = cString & " order by date,ID desc"
                    aRet = GetFields(cString)
                    If Not IsEmpty(aRet) Then
                        If IsDate(xDate.Text) Then
                            If grid1.TextMatrix(i, 9) <> "" Then
                                Dim nDays As Long
                                grid1.TextMatrix(i, 10) = Format(retFlag(aRet, "date"), "yyyy-mm-dd")
                                nDays = DateDiff("d", Format(retFlag(aRet, "date"), "yyyy-mm-dd"), Format(xDate.Text, "yyyy-mm-dd"))
                                grid1.TextMatrix(i, 11) = Round(nDays / 30, 2)
                                If Val(grid1.TextMatrix(i, 11)) < Val(grid1.TextMatrix(i, 9)) Then grid1.Cell(flexcpBackColor, i, 11, i, 11) = &H8080FF
                            End If
                            If grid1.TextMatrix(i, 15) <> "" Then
                                grid1.TextMatrix(i, 13) = Myvalue(retFlag(aRet, "COUNTER"))
                                grid1.TextMatrix(i, 14) = Myvalue(Val(grid1.TextMatrix(i, 12)) - Val(grid1.TextMatrix(i, 13)))
                                If Val(grid1.TextMatrix(i, 14)) < Val(grid1.TextMatrix(i, 15)) Then grid1.Cell(flexcpBackColor, i, 14, i, 14) = &H8080FF
                            End If
                        End If
                        
                    End If
                End If
            End If
        Next
    End If
End With
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cFileHeader
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
cString = cString & " Order by " & cFileHeader & ".DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub MyAddItem()
With grid1
.AddItem ""
If cdef_Box <> "" Then .TextMatrix(.Rows - 1, 0) = cdef_Box
If grid1.Rows > 2 Then
    .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
'    .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 2, 3)
End If
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .myRemove OldRow
        Calctotals
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Trim(.TextMatrix(nRow, 0)) = "" Then Exit Function
If Trim(.TextMatrix(nRow, 4)) = "" Then Exit Function
If (Not IsNumeric(.TextMatrix(nRow, 1))) Then Exit Function
If Trim(.TextMatrix(nRow, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 6 Then
    If Col >= 5 And Col < 10 Then
        grid1.Col = 12
    Else
        grid1.Col = Col + 1 + IIf(Col = 1 Or Col = 6, 1, 0)
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 4)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For i = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(Row, i)) = "" Then
        NextEmpty = i
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
Calctotals
End Sub
