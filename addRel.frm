VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form addrelfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«÷«›… «ﬁ«—»"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
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
   ScaleHeight     =   6150
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "«÷«›… ·«Ê· „—…"
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
      Height          =   285
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5490
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5400
      Width           =   4200
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   2025
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "addRel.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "addRel.frx":2008
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "addRel.frx":3FBF
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   5085
      Top             =   6300
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   741
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
      Bindings        =   "addRel.frx":62E2
      Height          =   5235
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   9825
      _cx             =   17330
      _cy             =   9234
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      TabBehavior     =   0
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
Attribute VB_Name = "addrelfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCode As String, pYear_code As String, pType As String
Dim con As New ADODB.Connection
Private Sub cmdAdd_Click()
Dim I As Long, cWhere As String
For I = 1 To grid1.rows - 1
    If mRound(grid1.TextMatrix(I, grid1.Cols - 2)) <> 0 Then
        cWhere = cWhere & turn(cWhere, ",") & grid1.TextMatrix(I, 0)
    End If
Next
cWhere_rel = cWhere
Unload Me
End Sub

Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
myloadgrd
End Sub
Private Sub myloadgrd()
atype = Claim_Type_Load(pType, "RELATION", con)
aYear = Year_Load(pYear_code, , con)
sDate1 = myFormat(retFlag(aYear, "date1"))
sdate2 = myFormat(retFlag(aYear, "date2"))

Dim cString As String
cString = "select file1_11.code,file1_11.Desca" & _
          ", relation_codes.Desca as relDesca,CONVERT(VARCHAR(10),File1_11.Date_Birth,111)," & AgeFieldRel("FILE1_11", sDate1, sdate2) & ", Cast(1 as Bit),file1_11.Relation from file1_11 inner join relation_codes on file1_11.relation = relation_codes.code " & _
          " where file1_11.member = " & sCode
If Not IsEmpty(atype) Then cString = cString & turn(cString) & "FILE1_11.RELATION  IN (" & atype & ")"
cString = cString & turn(cString) & "file1_11.date_begin >= " & DateSq(sDate1)
Set data1.Recordset = myRecordSet(cString, con)
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
.TextMatrix(0, 0) = "—ﬁ„ «· «»⁄"
.TextMatrix(0, 1) = "«·«”„"
.TextMatrix(0, 2) = "‰Ê⁄ «· »⁄Ì…"
.TextMatrix(0, 3) = " «—ÌŒ «·„Ì·«œ"
.TextMatrix(0, 4) = "«·”‰"
.TextMatrix(0, 5) = "«Œ Ì«—"

.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 1300
.ColWidth(4) = 1000
.ColWidth(5) = 1000
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColHidden(.Cols - 1) = True
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set addrelfrm = Nothing
End Sub
Private Sub Grid1_EnterCell()
grid1.Editable = grid1.Col = grid1.Cols - 2
End Sub

Private Sub Grid1_GotFocus()
Grid1_EnterCell
End Sub
