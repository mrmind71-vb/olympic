VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form BankInOutfrm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· Õ—ﬂ… «·»‰ﬂ"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14865
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
   ScaleHeight     =   8835
   ScaleWidth      =   14865
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   90
      Width           =   5490
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·„” ‰œ"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
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
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
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
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
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
         Left            =   4095
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   7065
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   7665
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
         Height          =   360
         Left            =   5175
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1290
      End
      Begin MSDataListLib.DataCombo XBANK 
         Height          =   315
         Left            =   2790
         TabIndex        =   7
         Top             =   630
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»‰ﬂ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6555
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   1515
      Begin VB.CommandButton CmdUndo 
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
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1365
      End
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
   Begin VB.Frame Frame8 
      Height          =   570
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7965
      Width           =   1920
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   8535
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   529
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
            TextSave        =   "09:04 „"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   465
      Left            =   2025
      Top             =   1170
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
      Height          =   6210
      Left            =   180
      TabIndex        =   19
      Top             =   1755
      Width           =   14550
      _cx             =   25665
      _cy             =   10954
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
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
End
Attribute VB_Name = "BankInOutfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim cStrBox As String, SearchItems As New Search3
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(1, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "Bank"
aInsert(1, 1) = addstring(XBANK.BoundText)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE5_30H", "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE5_30H")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_30H", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(7, 1)
With grid1
    For I = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
                
        aInsert(1, 0) = "code"
        aInsert(1, 1) = addvalue(grid1.TextMatrix(I, 0))
        
        aInsert(2, 0) = "date"
        aInsert(2, 1) = addDate(.TextMatrix(I, 2))
                
        aInsert(3, 0) = "Box"
        aInsert(3, 1) = addstring(.TextMatrix(I, 3))
                        
        aInsert(4, 0) = "Desca"
        aInsert(4, 1) = addstring(grid1.TextMatrix(I, 4))
        
        aInsert(5, 0) = "[value1]"
        aInsert(5, 1) = Val(grid1.TextMatrix(I, 5))

        aInsert(6, 0) = "[value2]"
        aInsert(6, 1) = Val(grid1.TextMatrix(I, 6))

        aInsert(7, 0) = "row"
        aInsert(7, 1) = I
        
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE5_30")
        Else
            con.Execute CreateUpdate(aInsert, "FILE5_30", " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = SearchItems.grid1.TextMatrix(SearchItems.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 1) = SearchItems.grid1.TextMatrix(SearchItems.grid1.Row, 1)
        If grid1.Row = grid1.Rows - 1 Then
            grid1.AddItem ""
            grid1.Select grid1.Row + 1, 0
        End If
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE5_30 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete From FILE5_30H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.EOF Then CardTable.MoveFirst
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
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE5_30H.Doc_No, CONVERT(VARCHAR(10),MIN(FILE5_30.[Date]),111),File5_10.Desca " & _
                  " FROM (FILE5_30H INNER JOIN FILE5_30 ON FILE5_30H.DOC_NO = FILE5_30.DOC_NO) INNER JOIN FILE5_10 ON FILE5_30H.BANK = FILE5_10.CODE"
Generalarray(2) = " group by FILE5_30H.Doc_No,FILE5_30.Date,File5_10.Desca order by FILE5_30H.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-»‰ﬂ-»‰œ «·Õ—ﬂ…"
listarray(0, 1) = "##FILE5_30.Date## " & _
                  " or ( %%file5_10.Desca%% )" & _
                  " Or (FILE5_30.code in (Select FILE5_30.code From File5_30 inner join file5_00 on file5_30.code = file5_00.code where %%FILE5_00.desca%%))"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "≈”„ «·»‰ﬂ"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "Customers Query"
Search3.Show 1
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
xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1, 6)
mydefine
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If xDoc_No.Enabled Then
    CmdNewInv_Click
Else
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub
Private Sub CmdUndo_Click()
If CardTable.RecordCount = 0 Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then
    CardTable.MoveLast
    myload
Else
    myload
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
cStrBox = StrBox
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT DOC_NO,BANK FROM FILE5_30H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "FILE5_10"

Set grid1.DataSource = data10
data10.ConnectionString = strCon

Set XBANK.RowSource = data1
XBANK.ListField = "Desca"
XBANK.BoundColumn = "CODE"

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    FixGrd
    mydefine
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
On Error Resume Next
Unload SearchItems
Set SearchItems = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    'grid1.TextMatrix(Row, 1) = RetSeek("file5_00", "ndxcode", Trim(grid1.TextMatrix(Row, 0)), "Desca")
    grid1.TextMatrix(Row, 1) = GetDesca("select Desca from file5_00 where code = " & MyParn(grid1.TextMatrix(Row, 0)))
End If
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 1 Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 5 Then
    grid1.Editable = IIf(Val(grid1.TextMatrix(grid1.Row, 6)) = 0, flexEDKbdMouse, flexEDNone)
ElseIf grid1.Col = 6 Then
    grid1.Editable = IIf(Val(grid1.TextMatrix(grid1.Row, 5)) = 0, flexEDKbdMouse, flexEDNone)
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And Col = 0 Then BankItemsLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE5_30 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 0 Then BankItemsLookup
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    If grid1.Rows > 3 Then
        grid1.TextMatrix(grid1.Rows - 2, 2) = grid1.TextMatrix(grid1.Rows - 3, 2)
        grid1.TextMatrix(grid1.Rows - 2, 3) = grid1.TextMatrix(grid1.Rows - 3, 3)
    End If
End If
End Sub
Private Function MYVALID() As Boolean
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF And xDoc_No.Enabled Then
    MsgBox "›« Ê—… »‰›” «·—ﬁ„ „‰ ﬁ»·"
    Exit Function
End If

If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If XBANK.BoundText = "" Then
    MsgBox "«·»‰ﬂ €Ì— „”Ã·"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "»‰Êœ €Ì— „”Ã·…"
    Exit Function
End If


With grid1
For I = 1 To .Rows - 2
'    If Trim(.TextMatrix(I, 2)) = "" Then
'        MsgBox "«·Œ“«‰… €Ì— „ÊÃÊœ…"
'        Exit Function
'    End If

'    If .TextMatrix(I, 0) = "" Then
'        .Select I, 0, I, grid1.Cols - 1
'        MsgBox "ﬂÊœ «·»‰œ €Ì— „ÊÃÊœ"
'        Exit Function
'    Else
If .TextMatrix(I, 0) <> "" Then
    If GetDesca("SELECT CODE FROM FILE5_00 WHERE CODE = " & MyParn(.TextMatrix(I, 0))) = "" Then
        .Select I, 0, I, 2
        MsgBox "ﬂÊœ «·»‰œ €Ì— ”·Ì„"
        Exit Function
    End If
End If
    If Val(.TextMatrix(I, 5)) = 0 And Val(.TextMatrix(I, 6)) = 0 Then
        MsgBox "ﬁÌ„… «·»‰œ €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
XBANK.BoundText = CardTable!bank
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE5_30.Code,file5_00.Desca,convert(varchar(10),DATE,111),BOX,FILE5_30.DESCA,VALUE1,VALUE2,ID " & _
               " FROM FILE5_30 LEFT JOIN FILE5_00 ON FILE5_30.CODE = FILE5_00.CODE " & _
               " where Doc_no = " & MyParn(xDoc_No.Text) & " Order by Row"

    data10.RecordSource = cString
    data10.Refresh
    grid1.AddItem ""
End With
FixGrd
End Sub
Private Sub mydefine()
XBANK.BoundText = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = (nMode = LoadMode And bEdit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
cmdSave.Enabled = bEdit
CmdDelInv.Enabled = bEdit
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!Code & ";" & boxtable!desca
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub BankItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE5_00"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
SearchItems.Caption = "≈” ⁄·«„ "
SearchItems.Show 1
End Sub
Private Sub FixGrd()
With grid1
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
     .FormatString = "«·Õ—ﬂ…|" & "«·Õ—ﬂ…|" & "«· «—ÌŒ|" & "Œ“‰…|" & "«·»Ì«‰|" & "«Ìœ«⁄« |" & "„”ÕÊ»« |"
    .ColWidth(0) = 500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 6000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColHidden(.Cols - 1) = True
    For I = 1 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    .ColComboList(3) = cStrBox
End With
End Sub
