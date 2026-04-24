VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form paymentfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "”œ«œ ›Ê« Ì— ¬Ã·"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Al-Hadith2"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   855
      Width           =   4785
      Begin VB.CheckBox xPrintCancel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·€«¡ «·ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   2130
      End
      Begin VB.TextBox xCopies 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ « ·‰”Œ :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   405
         Width           =   1680
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4905
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»«ﬁÌ :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   21.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì ”œ«œ ¬Ã·  :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   1860
      End
      Begin VB.Label xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label xTotalNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "’«›Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1410
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   8550
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   3300
      Begin VB.Label LblChange 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”œ«œ ‰ﬁœÌ :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·⁄—»Ê‰ :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Al-Hadith2"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label xCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label xAdvance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   675
         Width           =   1410
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1410
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4590
      Top             =   10530
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1710
      Top             =   360
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   3735
      Top             =   270
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   900
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   945
      OleObjectBlob   =   "payment.frx":0000
      Top             =   180
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3075
      Left            =   90
      TabIndex        =   8
      Top             =   1845
      Width           =   11715
      _cx             =   20664
      _cy             =   5424
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   14.25
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
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
   Begin Threed.SSCommand cmdExit 
      Height          =   1365
      Left            =   90
      TabIndex        =   9
      Top             =   4905
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   2408
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Al-Hadith2"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Œ—ÊÃ"
      ButtonStyle     =   2
      PictureAlignment=   11
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   1365
      Left            =   7875
      TabIndex        =   10
      Top             =   4905
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   2408
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Al-Hadith2"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Õ›Ÿ"
      ButtonStyle     =   2
      PictureAlignment=   11
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   1365
      Left            =   4005
      TabIndex        =   11
      Top             =   4905
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   2408
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Al-Hadith2"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄…"
      ButtonStyle     =   2
      PictureAlignment=   11
      BevelWidth      =   1
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   4545
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "paymentfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_No As String, nTotalInv As Single, nCash As Single, nAdvance As Single, sBoxName As String
Public sName As String, sCode As String
Public nValue As Single
Dim con As New adodb.Connection
Private Sub cmdExit_Click()
Unload Me
Set Reservefrm = Nothing
End Sub
Private Sub cmdSave_Click()
If myReplacegrd Then
    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
    For i = 0 To grid1.Rows - 1
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            If Me.xPrintCancel = 0 Then
                doprint
                Exit For
            End If
        End If
    Next
    Unload Me
End If
End Sub
Private Sub cmdPrint_Click()
If myReplacegrd Then
    myload
    doprint
Else
    MsgBox "·„ Ì „ «·Õ›Ÿ »‰Ã«Õ !! «·ÿ»«⁄… ·‰   „"
End If
End Sub


Private Sub Form_Load()
myloadsetting
MyloadCommand
Set grid1.DataSource = data1
openCon con
data1.ConnectionString = strCon
sBoxName = GetDesca("Select desca from file0_50 where code = " & MyParn(sboxSales))

xTotal.Caption = nTotalInv
xAdvance.Caption = nAdvance
xCash.Caption = nCash
xTotalNet.Caption = Val(xTotal.Caption) - Val(xAdvance.Caption) - Val(xCash.Caption)
myload
End Sub
Private Sub myload()
Dim cString
On Error GoTo myerror
cString = "SELECT  Convert(VARCHAR(10),INV_PAY.[DATE],111),FILE0_50.DESCA,INV_PAY.[VALUE],'' as balance,INV_PAY.BOX,ID " & _
          " FROM  INV_PAY INNER JOIN FILE0_50 ON INV_PAY.BOX = FILE0_50.code"
cString = cString & turn(cString) & " INV_PAY.[DOC_NO] = " & MyParn(sDoc_No)
cString = cString & " Order by INV_PAY.DATE,INV_PAY.ID"

data1.RecordSource = cString
data1.Refresh
With grid1
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = Format(sysDate, "yyyy/mm/dd")
    .TextMatrix(.Rows - 1, 1) = sBoxName
    .TextMatrix(.Rows - 1, 2) = Myvalue(nValue)
    .TextMatrix(.Rows - 1, 4) = sboxSales
End With
Calctotals
Fixgrd
grid1.Row = grid1.Rows - 1
grid1.Col = 2
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 6
.Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = "Al-Hadith2"
.Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.FormatString = "«· «—ÌŒ|«·ﬂ«‘Ì—|«·„”œœ|«·»«ﬁÌ|«·ﬂ«‘Ì—|"
.ColWidth(0) = 2500
.ColWidth(1) = 4000
.ColWidth(2) = 1200
.ColWidth(3) = 1200
.ColHidden(.Cols - 1) = True
For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Calctotals()
Dim nTotal As Single, nBalance As Single
With grid1
nBalance = nTotalInv - Val(xAdvance.Caption) - Val(xCash.Caption)
For i = 1 To grid1.Rows - 1
    nTotal = Round(nTotal + Val(.TextMatrix(i, 2)), 2)
    nBalance = Round(nBalance - Val(.TextMatrix(i, 2)), 2)
    .TextMatrix(i, 3) = nBalance
Next
xBalance.Caption = nBalance
xLate.Caption = nTotal
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set paymentfrm = Nothing
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        Calctotals
    End If
End If
End With
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 2 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" And KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”œ«œ „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from INV_PAY where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        Calctotals
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
    Calctotals
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Val(.TextMatrix(nRow, 2)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = Format(sysDate, "yyyy/mm/dd")
    .TextMatrix(.Rows - 1, 1) = sBoxName
    .TextMatrix(.Rows - 1, 4) = sboxSales
End If
Calctotals
End With
End Sub
Private Function myReplacegrd() As Boolean
Dim nCost As Double
Dim aInsert(4, 1)
With grid1
    On Error GoTo myerror
    con.BeginTrans
    For i = 1 To .Rows - 2
        If .TextMatrix(i, .Cols - 1) = "" Then
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(sDoc_No)
            
            aInsert(1, 0) = "[date]"
            aInsert(1, 1) = DateSq(grid1.TextMatrix(i, 0))
            
            aInsert(2, 0) = "BOX"
            aInsert(2, 1) = addstring(sboxSales)
    
            aInsert(3, 0) = "VALUE"
            aInsert(3, 1) = Val(.TextMatrix(i, 2))
        
            aInsert(4, 0) = "[SESSION]"
            aInsert(4, 1) = addvalue(sSession)
        
            con.Execute CreateInsert(aInsert, "INV_PAY")
        End If
    Next
    con.CommitTrans
End With
myReplacegrd = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Val(grid1.EditText) > Val(Val(grid1.TextMatrix(Row, 3))) Then
    Cancel = True
End If
End Sub
Private Sub MyloadCommand()
Dim cPath As String
cPath = App.Path & "\sys_images\menu"
On Error Resume Next
cmdSave.Picture = LoadPicture(cPath & "\" & cmdSave.Name & ".jpg")
cmdSave.PictureDisabled = LoadPicture(cPath & "\" & cmdSave.Name & "_Disabled.jpg")

CmdExit.Picture = LoadPicture(cPath & "\" & CmdExit.Name & ".jpg")
CmdExit.PictureDisabled = LoadPicture(cPath & "\" & CmdExit.Name & "_Disabled.jpg")

cmdPrint.Picture = LoadPicture(cPath & "\" & cmdPrint.Name & ".jpg")
cmdPrint.PictureDisabled = LoadPicture(cPath & "\" & cmdPrint.Name & "_Disabled.jpg")
If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub xCopies_Change()
addSetting "copies", Val(xCopies.Text), tempPath & "\payment.txt"
End Sub
Private Sub myloadsetting()
xPrintCancel.Value = Val(RetSetting("cancel", tempPath & "\payment.txt"))
xCopies.Text = IIf(Val(RetSetting("copies", tempPath & "\payment.txt")) = 0, 1, Val(RetSetting("copies", tempPath & "\payment.txt")))
End Sub
Private Function doprint() As Boolean
Dim sBoxName As String
sBoxName = GetDesca("Select desca from file0_50 where code = " & MyParn(sboxSales))
'On Error GoTo myerror
Dim temptable As New adodb.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

  
With grid1
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str1 = " ›’Ì·Ì ”œ«œ «·›« Ê—… —ﬁ„ : " & Format(sDoc_No)
    temptable!str2 = ArbString("«·⁄„Ì· : " & sName & "(" & Format(sCode) & ")")
    temptable!Str3 = "«·ﬂ«‘Ì— : " & sBoxName
    temptable!str4 = "«· «—ÌŒ : " & Format(sysDate, "yyyy/mm/dd")
    
    temptable!val4 = Val(xTotal.Caption)
    temptable!VAL5 = Val(xAdvance.Caption)
    temptable!val6 = Val(xCash.Caption)
    temptable!Val7 = Val(xTotalNet.Caption)
    temptable!Val8 = Val(xLate.Caption)
    temptable!val9 = Val(xBalance.Caption)
    If Format(sysDate, "yyyy/mm/dd") <> .TextMatrix(i, 0) Or .TextMatrix(i, 4) <> sboxSales Then
        temptable!Val11 = 0
        temptable!Str11 = ArbString("”œ«œ«  ”«»ﬁ… :")
    Else
        temptable!Val11 = 1
        temptable!Str11 = "”œ«œ «·ÌÊ„ ··ﬂ«‘Ì— : " & sBoxName
    End If
    temptable!str6 = TurnValue(.TextMatrix(i, 0))
    temptable!str7 = TurnValue(.TextMatrix(i, 1))
    temptable!val1 = Val(.TextMatrix(i, 2))
    temptable!val2 = Val(.TextMatrix(i, 3))
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
REPORT1.Reset
FixPrinter REPORT1
REPORT1.PrinterCopies = Val(xCopies.Text)
REPORT1.Destination = crptToPrinter
REPORT1.ReportFileName = App.Path & "\Reports\payment.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.Action = 1
'main.REPORT1.Destination = crptToWindow
doprint = True
closeCon:
temptable.Close
Set temptable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function
Private Function doprintCancel() As Boolean
Dim sBoxName As String
sBoxName = GetDesca("Select desca from file0_50 where code = " & MyParn(sboxSales))
'On Error GoTo myerror
Dim temptable As New adodb.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
 
With grid1
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str1 = " ›’Ì·Ì «” —Ã«⁄ ⁄—»Ê‰ ÕÃ“ —ﬁ„ : " & Format(sDoc_No)
    temptable!str2 = ArbString("«·⁄„Ì· : " & sName & "(" & Format(sCode) & ")")
    temptable!Str3 = "«·ﬂ«‘Ì— : " & sBoxName
    temptable!str4 = "«· «—ÌŒ : " & Format(sysDate, "yyyy/mm/dd")
    
    temptable!val4 = Val(xTotal.Caption)
    temptable!VAL5 = Val(xAdvance.Caption)
    temptable!val6 = Val(xCash.Caption)
    temptable!Val7 = Val(xTotalNet.Caption)
    temptable!Val8 = Val(xLate.Caption)
    temptable!val9 = Val(xBalance.Caption)
    If Format(sysDate, "yyyy/mm/dd") <> .TextMatrix(i, 0) Or .TextMatrix(i, 4) <> sboxSales Then
        temptable!Val11 = 0
        temptable!Str11 = ArbString("”œ«œ«  ”«»ﬁ… :")
    Else
        temptable!Val11 = 1
        temptable!Str11 = "”œ«œ «·ÌÊ„ ··ﬂ«‘Ì— : " & sBoxName
    End If
    temptable!str6 = TurnValue(.TextMatrix(i, 0))
    temptable!str7 = TurnValue(.TextMatrix(i, 1))
    temptable!val1 = Val(.TextMatrix(i, 2))
    temptable!val2 = Val(.TextMatrix(i, 3))
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.PrinterCopies = Val(xCopies.Text)
main.REPORT1.Destination = crptToPrinter
main.REPORT1.ReportFileName = App.Path & "\Reports\payment.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
'main.REPORT1.Destination = crptToWindow
doprintCancel = True
closeCon:
temptable.Close
Set temptable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function


