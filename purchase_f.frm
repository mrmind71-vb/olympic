VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Purchasefrm_f 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
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
   ScaleHeight     =   9765
   ScaleWidth      =   15195
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   12375
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   2850
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
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   630
      Width           =   10770
      Begin VB.TextBox xCode 
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
         Left            =   8460
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   1095
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
         Left            =   8460
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1095
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
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   7110
         TabIndex        =   24
         Top             =   855
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
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
         Index           =   1
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   945
         Width           =   585
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   495
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   495
         Width           =   3165
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   600
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
         Left            =   9645
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Index           =   0
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   9645
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   570
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6720
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1845
      Width           =   15000
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   6495
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   14820
         _cx             =   26141
         _cy             =   11456
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
   Begin VB.Frame Frame8 
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   8505
      Width           =   1905
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
         TabIndex        =   21
         ToolTipText     =   "Move Last"
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   135
         Width           =   435
      End
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
         TabIndex        =   18
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2745
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -1260
      Top             =   990
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
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
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   900
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   1755
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   630
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
End
Attribute VB_Name = "Purchasefrm_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset, ClientTable As ADODB.Recordset, cFileHeader As String, rdPaid As New ADODB.Recordset, bNoMsgExit As Boolean
Dim ItemTable As New ADODB.Recordset
Dim SEARCH31 As New Search3, search32 As New Search3, bMarket As Boolean
Dim dLastdate As String, bEdit As Boolean, cWhereType As String
Dim cFile As String, cFileClient, cMoveName, cFileMove, cFileSerial, cItemmove As String, cClientmove, cFieldItem, cFieldClient, cCodeDesca As String
Dim defBox As String
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.cost From file1_10 left join file1_50 on file1_10.group = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«· ﬂ·›…"
GrdArray(3, 1) = 0


searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
        
    grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    grid1.TextMatrix(grid1.Row, 3) = "1"
    GrdDesc grid1.Row
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = 1
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
        MakeSerial
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 3) = 1
        grid1.Select grid1.Rows - 1, 1
    End If
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    CalcTotals
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(SEARCH31.grid1.TextMatrix(SEARCH31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    SEARCH31.Hide
    MyLoad
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    Unload search32
End If
Exit Sub
myerror:
End Sub
Private Sub cmdBarCode_Click()
    Dim tBarCode As New ADODB.Recordset
    If grid1.Rows = 1 Then Exit Sub
    tBarCode.Open "addprint", CON, adOpenKeyset, adLockReadOnly, adCmdTable
    tBarCode.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If Not tBarCode.EOF Then
        If MsgBox("«·›« Ê—…  „  —ÕÌ·Â« „‰ ﬁ»· .. «·€«¡ «·«’‰«› «·„—Õ·… ··ÿ»«⁄…", vbOKCancel + vbDefaultButton2, " —ÕÌ· ··ÿ»«⁄…") = vbCancel Then
            Exit Sub
        End If
    End If
    With grid1
    CON.Execute "DELETE * FROM ADDPRINT WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    For I = 1 To grid1.Rows - 2
        If Val(retitem(.TextMatrix(I, 1), "TYPE") & "") = "0" Then
            CON.Execute "Insert Into ADDPRINT(Doc_no,Item,Quant,isPrint) " & _
                        " Values(" & _
                        addstring(xDoc_No.Text) & "," & _
                        addstring(.TextMatrix(I, 1)) & "," & _
                        addvalue(.TextMatrix(I, 3)) & "," & _
                        "TRUE" & _
                        ")"
        End If
    Next
    End With
End Sub

Private Sub cmdClient_Click()
publicFlag = 2
Clients.Show 1
End Sub
Private Sub CmdExit_Click()
If Not bNoMsgExit Then If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub CmdNewInv_Click()
myDefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
'If myPublic = 0 Then If Not nofoundOther Then Exit Sub
mySave
End Sub
Private Sub CmdUndo_Click()
If MsgBox(" —«Ã⁄ ⁄‰  ”ÃÌ· «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
MyLoad
End Sub
Private Sub Cmditem_Click()
Dim bEditLocal As Boolean
bEditLocal = bEdit: bEdit = True
items.Show 1
bEdit = bEditLocal
End Sub

Private Sub Command1_Click()
mySave False
Dream_Bar.Show 1
End Sub
Private Sub Command3_Click()
mySave False
doprint
End Sub
Private Sub Command4_Click()
Dim locTable As New ADODB.Recordset
locTable.Open "FILE1_10", CON, adOpenKeyset, adLockReadOnly, adCmdTable
On Error GoTo myerror
CON.BeginTrans
If Not (locTable.EOF And locTable.BOF) Then
    locTable.MoveLast
    nRecordCount = locTable.RecordCount
    locTable.MoveFirst
End If
I = 0
prog1.Value = 0
prog1.Visible = True
Do Until locTable.EOF
    I = I + 1
    prog1.Value = Round(I / nRecordCount, 2) * 100
    CON.Execute "UPDATE FILE1_10 SET FILE1_10.COST = " & itemCost(locTable!Item) & " where item = " & MyParn(locTable!Item)
    locTable.MoveNext
Loop
'CON.Execute "update (file1_10 inner join file7_20 on file1_10.item = file7_20.item) inner join file7_20h on file7_20.doc_no = file7_20h.doc_no set file1_10.supler = file7_20h.code"
CON.CommitTrans
prog1.Visible = False
MsgBox "DONE..."
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
CON.RollbackTrans
End Sub

Private Sub Command5_Click()
mySave False
doprint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bEdit = True
Select Case myPublic
Case 0
    cCodeDesca = "«·„Ê—œ"
    cFile = "File7_20"
    cFileHeader = "File7_20H"
    cFileClient = "File4_10"
    cFileMove = "File4_11"
    cFieldItem = "[IN]"
    cFieldClient = "[SAL]"
    cMoveName = "„‘ —Ì« "
    cItemmove = 2
    cClientmove = 2
    lblClient.Caption = "«·„Ê—œ :"
    Me.Caption = "›« Ê—… „‘ —Ì« "
    cmdClient.Caption = "„Ê—œ ÃœÌœ"
    cWhereType = " WHERE [TYPE] = 1"
Case 1
    cCodeDesca = "«·„Ê—œ"
    cFile = "FILE7_30"
    cFileHeader = "File7_30H"
    cFileClient = "File4_10"
    cFileMove = "File4_11"
    cFieldItem = "[OUT]"
    cFieldClient = "[PAY]"
    cMoveName = "„—œÊœ „‘ —Ì« "
    cItemmove = 4
    cClientmove = 3
    lblClient.Caption = "«·„Ê—œ :"
    Me.Caption = "›« Ê—… „—œÊœ „‘ —Ì« "
'    FramePaid.Visible = False
    'cmdClient.Caption = "„Ê—œ ÃœÌœ"
    'cmdClient.Enabled = RetSec("tmSupData")
End Select
'cmditem.Enabled = RetSec("tmItem")
ItemTable.Open "file1_10", CON, adOpenStatic, adLockReadOnly, adCmdTable


Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM " & cFileHeader & cWhereType & " ORDER BY DOC_NO", CON, adOpenKeyset, adLockReadOnly, adCmdText

Set ClientTable = New ADODB.Recordset
ClientTable.Open cFileClient, CON, adOpenKeyset, adLockReadOnly, adCmdTable

data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

data2.ConnectionString = CON.ConnectionString
data2.RecordSource = "SELECT * FROM FILE0_50"
Set xBox.RowSource = data2
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

defBox = retDef("file0_50")
With grid1
    .Cols = 9
    .Rows = 2
End With

Set grid1.DataSource = DATA3
DATA3.ConnectionString = CON.ConnectionString

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    myDefine
    FixGrd
    xDoc_No.Text = RetZero("1", 6)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ItemTable.Close
Set ItemTable = Nothing
Unload Search3
Unload SEARCH31
Unload search32
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 1 Then
    'If Row <> 1 Then grid1.TextMatrix(Row, 1) = myShortCut(Trim(grid1.TextMatrix(Row, 1)), Trim(grid1.TextMatrix(Row - 1, 1)))
    GrdDesc Row
End If
CalcTotals
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.Select 1, 1
End If

End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then
    grid1.AddItem "", grid1.Row
    MakeSerial grid1.Row - 1
End If
If KeyCode = 112 Then
    If grid1.Col = 1 And grid1.Row <> 0 Then ItemsLookup
End If

If KeyCode = 112 Then
    If grid1.Col = 1 And grid1.Row <> 0 Then ItemsLookup
'    If grid1.Col <> 1 And grid1.Row <> 0 Then grid1.Cell(flexcpBackColor, grid1.Row, 1, grid1.Row, grid1.Cols - 1) = IIf(grid1.Cell(flexcpBackColor, grid1.Row, 1, grid1.Row, grid1.Cols - 1) = &H8000000F, &H80000005, &H8000000F)
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    MakeSerial
'    If Row <> 1 Then grid1.TextMatrix(Row, 4) = grid1.TextMatrix(Row - 1, 4)
End If
End Sub

Private Sub XBOX_Click(Area As Integer)
If Not xBox.MatchedWithList Then xBox.BoundText = "'"
End Sub

Private Sub xCode_DblClick()
CLIENTLOOKUP
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
xCodeDesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xCode.Text)) & ""
'xBalance.Caption = Format(GetDesca("Select sum(Format(val(SAL & '') - val (pay & ''),'Fixed')) FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
End Sub
Private Sub xDiscount_LostFocus()
CalcTotals
End Sub
Private Function MYVALID() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
For I = 1 To grid1.Rows - 2
    If Not validRow(I) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad(Optional bLeaveBal As Boolean = False)
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!Store
xCode.Text = CardTable!CODE & ""
xBox.BoundText = CardTable!Box & ""
xCodeDesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(xCode.Text)) & ""
xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = TurnValue(Val(CardTable!Discount & ""), 0, "")
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")

With grid1
    cString = "SELECT " & cFile & ".ROW, " & cFile & ".ITEM, FILE1_10.DESCA, Quant,Format(" & cFile & ".Price,'fixed'),0 as Total," & cFile & ".discount,FILE1_10.PACKAGE,FILE1_10.UNIT" & _
          " FROM " & cFile & " LEFT JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " order by " & cFile & ".ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
    MakeSerial
End With
Handlecontrols LoadMode
CalcTotals
FixGrd
End Sub
Private Sub myDefine()
If CardTable.EOF And CardTable.BOF Then
    xDoc_No.Text = RetZero("1")
Else
    xDoc_No.Text = RetZero(GetDesca("Select max(doc_no) from  " & cFile & cWhereType))
End If
xBox.BoundText = defBox
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = ""
xCodeDesca.Caption = ""
xBalance.Caption = ""
xCode.Text = ""
xDiscount.Text = ""
xtotal.Caption = ""
xTax.Text = ""
xTotalItem.Caption = ""
xTotalDis.Caption = ""
xTotalDisItem.Caption = ""
xDisItem.Caption = ""
xusername.Text = ""
xtotalQuant.Caption = ""
xRate.Text = ""
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
CmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad True
End Sub
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem grid1.Row
        CalcTotals
        MakeSerial grid1.Row
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 1
        If KeyCode = 27 Then
            Exit Sub
        End If
        If KeyCode = 112 Then
            ItemsLookup
        End If
End Select
End Sub
Private Sub GrdDesc(Row)
'Grid1.TextMatrix(Row, 2) = ""
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
ItemTable.Find "item = " & MyParn(grid1.TextMatrix(Row, 1)), , adSearchForward, adBookmarkFirst
If Not ItemTable.EOF Then
    grid1.TextMatrix(Row, 2) = ItemTable!Desca
    grid1.TextMatrix(Row, 4) = LastPrice(grid1.TextMatrix(Row, 1))
    grid1.TextMatrix(Row, 7) = ItemTable!package & ""
    grid1.TextMatrix(Row, 8) = ItemTable!unit & ""
End If
End Sub
Private Function CalcTotals()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalDis As Double
With grid1
For I = 1 To grid1.Rows - 2
    nTotalitem = nTotalitem + (Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)))
    nDiscount = 1 - (Val(.TextMatrix(I, 6)) / 100)
    grid1.TextMatrix(I, 5) = Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount
    nTotalDisItem = nTotalDisItem + (Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount)
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 3))
Next
nDisItem = nTotalitem - nTotalDisItem
nTotalDis = nTotalDisItem - Val(xDiscount.Text)
nTotal = nTotalDis + Val(xTax.Text)
If nTotalitem <> 0 Then
    xRateDis.Text = TurnValue(Round((nTotalDisItem - nTotalDis) / nTotalDisItem * 100, 2), 0, "")
End If
xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
xTotalItem.Caption = Format(nTotalitem, "Fixed")
xTotalDisItem.Caption = Format(nTotalDisItem, "Fixed")
xDisItem.Caption = Format(nDisItem, "Fixed")
xTotalDis.Caption = Format(nTotalDis, "Fixed")
xtotal.Caption = Format(nTotal, "Fixed")
xtotalQuant.Caption = Format(nTotalQuant, "#0.0000")
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Format([DATE],'yyyy/mm/dd'), " & cFileClient & ".Desca " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE )"
                  cWhereType

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ " & cCodeDesca
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load SEARCH31
SEARCH31.Caption = "«” ⁄·«„"
SEARCH31.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow And Trim(grid1.TextMatrix(I, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function

Private Sub xRate_LostFocus()
If Val(xRate.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
End Sub
Private Function validRow(nRow) As Boolean
If nRow > 0 Then
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
    If Trim(grid1.TextMatrix(nRow, 2)) = "" Then Exit Function
   ' If Val(grid1.TextMatrix(nRow, 5)) = 0 Then Exit Function
End If
validRow = True
End Function
Sub additemProc()
grid1.RemoveItem grid1.Rows - 1
With additemfrm.grid1
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(I, 4)) <> 0 Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
            grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(I, 0)
            grid1.TextMatrix(grid1.Rows - 1, 2) = retitem(.TextMatrix(I, 0), "desca")
            grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(I, 4)
            grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(I, 5)
        End If
    Next
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    CalcTotals
End With
End Sub
Private Function RetItemBalance(citem, cStore, dDate) As Double
If citem = "" Then Exit Function
movetable.Seek Array(citem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> citem Or cStore <> movetable!Store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    'If Not (movetable!Type = cItemmove And movetable!Doc_Id = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    'End If
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub FixGrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ì|" & "«·Œ’„|" & "«·⁄»Ê…|" & "«·ÊÕœ…"
.ColWidth(0) = 500
.ColWidth(1) = 1800
.ColWidth(2) = 4500
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
    If I > 3 Then grid1.ColHidden(I) = True
Next
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From " & cFileClient
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub

Private Sub xTax_LostFocus()
CalcTotals
End Sub
Private Sub doprint()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str21 = "›« Ê—… „‘ —Ì«  —ﬁ„ : " & Format(xDoc_No.Text)
    temptable!str1 = TurnValue(xCodeDesca.Caption)
    temptable!str2 = TurnValue(xStore.Text)
    temptable!Str11 = TurnValue(xDate.Text)
    temptable!str3 = TurnValue(grid1.TextMatrix(I, 1))
    temptable!str4 = TurnValue(grid1.TextMatrix(I, 2))
    temptable!val2 = TurnValue(Val(grid1.TextMatrix(I, 3)))
    temptable!val1 = TurnValue(Val(grid1.TextMatrix(I, 4)))
    temptable!val3 = TurnValue(Val(grid1.TextMatrix(I, 6)))
    temptable!Val11 = Val(retitem(grid1.TextMatrix(I, 1), "Price") & "")
    temptable!Val10 = I
    temptable!val4 = Val(xTotalItem.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xTotalDis.Caption)
    temptable!Val7 = Val(xTax.Text)
    temptable!Val8 = Val(xtotal.Caption)
    temptable!Val10 = Val(grid1.TextMatrix(I, 8))
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\purchase.rpt"
Main.REPORT1.DataFiles(0) = tempPath
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub mySave(Optional bMsg As Boolean = True)
If Not MYVALID Then Exit Sub
CalcTotals
If Not MyReplace Then Exit Sub
CardTable.Requery
If bMsg Then MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
'xBalance.Caption = Format(GetDesca("Select sum(val(SAL & '') - val (pay & '')) as balance FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
MyLoad
End Sub
Sub myproc2(nDoc_no)
bNoMsgExit = True
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    MyLoad
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub

