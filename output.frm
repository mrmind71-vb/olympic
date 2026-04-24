VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form outputfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’«œ—"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19350
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
   ScaleHeight     =   9600
   ScaleWidth      =   19350
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
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1350
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   8685
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   495
      Width           =   1365
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
         Picture         =   "output.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "output.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   13860
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "output.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "output.frx":6CFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "output.frx":9594
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "output.frx":BB40
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   675
      Width           =   9195
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
         Height          =   315
         Left            =   6750
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1290
      End
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   5310
         TabIndex        =   5
         Top             =   540
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   210
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -300
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
      Height          =   6810
      Left            =   135
      TabIndex        =   15
      Top             =   1665
      Width           =   19140
      _cx             =   33761
      _cy             =   12012
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   9
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
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   15975
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   8505
      Width           =   3300
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   20
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
         Picture         =   "output.frx":E313
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "output.frx":104CF
      End
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   17
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
         Picture         =   "output.frx":1261E
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "output.frx":147EE
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   18
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
         Picture         =   "output.frx":16936
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "output.frx":18AFE
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   19
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
         Picture         =   "output.frx":1AC4D
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "output.frx":1CE2D
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   21
      Top             =   9210
      Width           =   19350
      _ExtentX        =   34131
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
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
End
Attribute VB_Name = "outputfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Integer
Dim CardTable As ADODB.Recordset
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchCar As New Search3
Dim cList As String, cList2 As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))
con.BeginTrans
On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE1_81H", "doc_no")))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "FILE1_81H")
Else
    con.Execute addUpdate(aInsert, "FILE1_81H", "doc_no = " & addstring(xDoc_No.Text))
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
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        Dim bNew As Boolean
        bNew = grid1.Row = grid1.Rows - 1
        grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = "1"
        GrdDesc grid1.Row
        Grid1_AfterEdit grid1.Row, grid1.Col
        If Not bNew Then
            Unload oSearchItem
            CellPos 13, grid1.Row, 2
        Else
            grid1.Select grid1.Rows - 1, 0
        End If
    ElseIf grid1.Col = 5 Then
        grid1.TextMatrix(grid1.Row, 5) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 6) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 4)
        grid1.TextMatrix(grid1.Row, 7) = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 1) & " " & oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 2)
        Unload oSearchCar
        Grid1_AfterEdit grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    oSearchDoc.Hide
    myUndo
ElseIf ActiveControl.Name = xDoc_No.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    oSearchDoc.Hide
Else
    ActiveControl.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Unload oSearchItem
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub chkPrv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fixgrd
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute " Delete  From FILE1_81 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute " Delete  From FILE1_81H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
       CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
       If CardTable.BOF Then CardTable.MoveFirst
       myload
    End If
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA " & _
                  " FROM FILE1_81H INNER JOIN FILE0_40 ON  Store = FILE0_40.CODE"

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·„Œ“‰"
GrdArray(2, 1) = 2000
searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
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
On Error Resume Next
xDoc_No.SetFocus
Err.Clear
End Sub

Private Sub CmdPrint_Click()
doprint
End Sub

Private Sub cmdSave_Click()
foundOther
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
If xDoc_No.Tag = LoadMode Then grid1.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub
Private Sub Form_Load()
bedit = True
openCon con
cList = StrList2("Select CODE,DESCA FROM DRIVER ORDER BY DESCA")
cList2 = StrList2("Select CODE,DESCA FROM FILE4_10 ORDER BY DESCA")
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon
openCardTable
myUndo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchItem
Unload oSearchDoc
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
tBalStore.Close
Set CardTable = Nothing
Set tBalStore = Nothing
closeCon con
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then
    Calctotals
    Exit Sub
End If
With grid1
If Row = grid1.Rows - 1 Then
    myAddItem
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
If grid1.Col = 13 Then
    Dim cWhere As String
    If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
    cWhere = "item = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
    If IsDate(xDate.Text) Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_81H.date <=" & DateSq(xDate.Text)
    cWhere = cWhere & turn(cWhere, " and ") & "car = " & grid1.TextMatrix(grid1.Row, 5)
    cWhere = cWhere & turn(cWhere, " and ") & "id <> " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
    showfrm2.cWhere = cWhere
    showfrm2.sDate = xDate.Text
'    showfrm2.sCaption = "„ «»⁄… «·„’—Ê› " & grid1.TextMatrix(grid1.Row, 1) & " ··”Ì«—… " & grid1.TextMatrix(grid1.Row, 7)
    showfrm2.Show 1
End If
End Sub

Private Sub Grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 2 Or grid1.Col = 5 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row < 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 8 And Col <> 9 And Col <> 10 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim aRet As Variant
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Cancel = True
    Else
        aRet = ItemFields(grid1.EditText, con)
        If IsEmpty(aRet) Then
           MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
           Cancel = True
        Else
            grid1.TextMatrix(Row, 1) = retFlag(aRet, "desca") & ""
            If Val(grid1.TextMatrix(Row, 3)) = 0 Or Trim(grid1.TextMatrix(grid1.Row, 0)) <> Trim(grid1.EditText) Then
                grid1.TextMatrix(Row, 3) = LastCostDate(grid1.EditText, Format(xDate.Text, "YYYY-MM-DD"), con)
            End If
        End If
    End If
ElseIf Col = 5 Then
    If (Not ValidInt(grid1.EditText)) Then
        If Trim(grid1.EditText) <> "" Then
            Cancel = True
        Else
            grid1.TextMatrix(grid1.Row, 6) = ""
            grid1.TextMatrix(grid1.Row, 7) = ""
        End If
    Else
        aRet = GetFields("SELECT BOARD,DESCA FROM CARS WHERE CODE = " & grid1.EditText)
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
Private Sub xdoc_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If Not bIgMsg Then
    With grid1
    For i = 1 To .Rows - 2
        If .TextMatrix(i, 0) = "" Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        ElseIf Not validItem(.TextMatrix(i, 0), con) Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
        If Val(.TextMatrix(i, 2)) = 0 Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
            Exit Function
        End If
    Next
    End With
End If
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!DOC_NO & ""
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xStore.BoundText = CardTable!store
myloadgrd
Handlecontrols LoadMode
Calctotals
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag("FILE1_81H", "doc_no")))
xDate.Text = Format(Date, "YYYY-MM-DD")
xStore.BoundText = ""
grid1.Rows = 1
grid1.AddItem ""
StatusBar1.Panels(1).Text = ""
Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And DOC_NO = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
ElseIf xDoc_No.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 112 And grid1.Col = 5 Then
    carsLookupAll Me, oSearchCar
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from FILE1_81 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 3) = ""
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub
Dim aRet As Variant
aRet = ItemFields(grid1.TextMatrix(Row, 0), con)
If Not IsEmpty(aRet) Then
    grid1.TextMatrix(Row, 1) = retFlag(aRet, "desca") & ""
    grid1.TextMatrix(Row, 3) = LastCostDate(grid1.TextMatrix(Row, 0), Format(xDate.Text, "YYYY-MM-DD"), con)
End If
End Sub
Private Function Calctotals()
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
For i = 1 To grid1.Rows - 2
    .TextMatrix(i, 4) = Val(grid1.TextMatrix(i, 2)) * Val(grid1.TextMatrix(i, 3))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 2))
    nTotalCost = nTotalCost + Val(grid1.TextMatrix(i, 4))
Next
StatusBar1.Panels(1).Text = turn(Myvalue(nTotalQuant), "≈Ã„«·Ì ⁄œœ «·«’‰«› : ") & Myvalue(nTotalQuant)
StatusBar1.Panels(2).Text = turn(Myvalue(nTotalQuant), "≈Ã„«·Ì  ﬂ·›… «·«’‰«› : ") & Myvalue(nTotalCost)
End With
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub foundOther()
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ ====> " & nRow
        Exit Sub
    End If
Next
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str21 = "„” ‰œ " & cTitle & " " & Format(xDoc_No.Text)
    temptable!date3 = DateFix(xDate.Text)
    temptable!str2 = TurnValue(xStore1.Text)
    temptable!str3 = TurnValue(xStore2.Text)
    temptable!str4 = TurnValue(grid1.TextMatrix(i, 0))
    temptable!str5 = TurnValue(grid1.TextMatrix(i, 1))
    temptable!val2 = TurnValue(Val(grid1.TextMatrix(i, 2)))
    temptable!val1 = Val(GetDesca("select price from file1_10 where item = " & MyParn(grid1.TextMatrix(i, 0))) & "")
    temptable!Val3 = Val(GetDesca("select DISCOUNT from file1_10 where item = " & MyParn(grid1.TextMatrix(i, 0))) & "")
    temptable!val10 = i
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\TRANS.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub Fixgrd()
With grid1
    .FormatString = "«·ﬂÊœ|" & "«·’‰›|" & "«·ﬂ„Ì…|" & "«· ﬂ·›…|" & "«·«Ã„«·Ì|" & "«·”Ì«—…|" & "—ﬁ„ «··ÊÕ…|" & "»Ì«‰ «·”Ì«—…|" & "«·„”∆Ê·|" & "Œ—ÊÃ »Ê«”ÿ…|" & "«·„Ê—œ|" & "„œ… «·«Â·«ﬂ »«·‘Â—|" & "«Œ—  «—ÌŒ «’·«Õ|" & "«·„œ… »«·‘Â—|"
    .RowHeight(0) = 600
    .ColWidth(0) = 1000
    .ColWidth(1) = 3800
    .ColWidth(2) = 800
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .ColWidth(6) = 1300
    .ColWidth(7) = 2500
    .ColWidth(8) = 2000
    .ColWidth(9) = 2000
    .ColWidth(10) = 2500
    .ColComboList(8) = cList
    .ColComboList(9) = cList
    .ColComboList(10) = cList2
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(10) = True
    
    .ColHidden(.Cols - 1) = True
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    If chkPrv.Value = 1 Then
        For i = 1 To grid1.Rows - 2
            If Trim(grid1.TextMatrix(i, 1)) <> "" And Trim(grid1.TextMatrix(i, 5)) <> "" And grid1.TextMatrix(i, .Cols - 1) <> "" Then
                Dim aRet As Variant
                cString = "select top 1 date from file1_81h inner join file1_81 on file1_81h.doc_no = file1_81.doc_no "
                cString = cString & turn(cString) & "car = " & grid1.TextMatrix(i, 5)
                cString = cString & turn(cString) & "ITEM = " & MyParn(grid1.TextMatrix(i, 0))
                cString = cString & turn(cString) & "date <= " & DateSq(xDate.Text)
                cString = cString & turn(cString) & "id <> " & grid1.TextMatrix(i, .Cols - 1)
                cString = cString & " order by date desc"
                aRet = GetFields(cString)
                If Not IsEmpty(aRet) Then
                    grid1.TextMatrix(i, .Cols - 3) = Format(retFlag(aRet, "date"), "yyyy-mm-dd")
                    If IsDate(xDate.Text) Then
                        Dim nDays As Long
                        nDays = DateDiff("d", Format(retFlag(aRet, "date"), "yyyy-mm-dd"), Format(xDate.Text, "yyyy-mm-dd"))
                        grid1.TextMatrix(i, .Cols - 2) = Round(nDays / 30, 2)
                        If Val(grid1.TextMatrix(i, .Cols - 2)) < Val(grid1.TextMatrix(i, grid1.Cols - 4)) Then grid1.Cell(flexcpBackColor, i, .Cols - 2, i, .Cols - 1) = &H8080FF
                    End If
                End If
            End If
        Next
    End If

End With
End Sub
Private Sub myreplaceGrd(Row As Long)
Dim aInsert
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(nRow = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "ITEM", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "QUANT", Val(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "COST", Val(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "car", addvalue(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "CODE2", addstring(grid1.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "CODE_SUP", addstring(grid1.TextMatrix(i, 10)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_81")
        Else
            con.Execute addUpdate(aInsert, "FILE1_81", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
cString = "Select FILE1_81.ITEM,FILE1_10.DESCA,FILE1_81.Quant,FILE1_81.COST,FILE1_81.TOTAL,FILE1_81.CAR,CARS.BOARD, CARS.DESCA ,FILE1_81.CODE,FILE1_81.CODE2,FILE1_81.CODE_SUP,FILE1_10.MONTHES ,'' AS DATE_PRV,'' AS PERIOD_PRV,FILE1_81.ID" & _
          " From FILE1_81 inner join file1_10 on FILE1_81.item = file1_10.item LEFT JOIN CARS ON FILE1_81.CAR = CARS.CODE"
cString = cString & turn(cString) & " FILE1_81.DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & " Order by FILE1_81.ID"
data10.RecordSource = cString
data10.Refresh
grid1.AddItem ""
Fixgrd
Calctotals
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 10 Then
    grid1.Col = Col + 1 + IIf(Col = 0, 1, 0) + IIf(Col = 2, 2, 0) + IIf(Col = 5, 2, 0)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 0, 2)
    grid1.ShowCell grid1.Row, 0
Else
    grid1.Select Row, Col
End If
End Sub
Private Function validRow(nRow) As Boolean
If Not MYVALID(True) Then Exit Function
If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
validRow = True
End Function
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem grid1.Row
End Sub
Private Sub myAddItem()
grid1.AddItem ""
End Sub
Private Sub openCardTable()
Set CardTable = New ADODB.Recordset
Dim cString As String
cString = "SELECT * FROM FILE1_81H"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xDoc_No.Text) <> "" Then
        CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
    grid1.Select grid1.Rows - 1, 0
    grid1.ShowCell grid1.Row - 1, 0
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End If
End Sub
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
