VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form grdHonorfrm 
   Caption         =   "»Ì«‰«  «·«⁄÷«¡ «·‘—›ÌÌ‰ «·Õ«·ÌÌ‰"
   ClientHeight    =   10110
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   18900
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
   ScaleWidth      =   18900
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   630
      Width           =   4965
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3735
         Picture         =   "grdHonor.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdHonor.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grdHonor.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   2430
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
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
         Picture         =   "grdHonor.frx":7149
         Caption         =   "ÿ»«⁄…"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdHonor.frx":A147
      End
   End
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   9735
      Width           =   18900
      _ExtentX        =   33338
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
      Height          =   1365
      Left            =   7830
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   9915
      Begin VB.CheckBox chkCurrent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Õ«·ÌÌ‰ ›ﬁÿ"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.TextBox xNo2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox xnotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2580
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox xNo1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   1005
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   900
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xSeason 
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Õ Ì —ﬁ„"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "„Ê”„ «·ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "—ﬁ„ ﬂ«—‰ÌÂ „‰"
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label19 
         Caption         =   " «—ÌŒ «‰ Â«¡ „‰ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8325
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
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
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label xCodeDesca 
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   2445
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„·ÕÊŸ…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8370
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   585
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2790
      Top             =   405
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1170
      Top             =   585
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
      Height          =   6270
      Left            =   45
      TabIndex        =   13
      Top             =   1395
      Width           =   17745
      _cx             =   31300
      _cy             =   11060
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
      Cols            =   8
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   18
      Top             =   9585
      Visible         =   0   'False
      Width           =   18900
      _ExtentX        =   33338
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "grdHonorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer
Dim oSearchMember As New Search, bNoCheck As Boolean
Dim con As New ADODB.Connection
Dim aHeader()
Private Sub Check1_Click()
If Not bNoCheck Then
    myload
End If
bNoCheck = False
End Sub
Private Sub cmdExel_Click()
ToFileExel2 grid1, , , , , 0.9
End Sub
Private Sub cmdPrint_Click()
Dim nRate As Double, i As Long
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", 3)
'aRow(0) = AddFlag(aRow(0), "text", "«·≈Ã„«·Ì")
Dim nwidth As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = grid1.ColWidth(i) + nwidth
    End If
Next


nRate = grdRate(grid1, 11000)
Set PrintGrdNew.myForm = Me

Me.MousePointer = 11
grid1.ColHidden(1) = True
PrintGrdNew.doPrint grid1, nRate, 0, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, False, 12, , aRow
grid1.ColHidden(1) = False
Me.MousePointer = 0
PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set grdpaid1 = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = 11
myload
Me.MousePointer = 0
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT TOP 2 CODE,DESCA FROM YEARS_CODES ORDER BY CODE DESC", con)
Set xSeason.RowSource = data1
xSeason.ListField = "Desca"
xSeason.BoundColumn = "Code"

Set grid1.DataSource = data10
grid1.ExplorerBar = flexExSortShow

Fixgrd
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
ReDim aHeader(4)
With grid1
cString = "SELECT FILE3_10.CODE, FILE3_10.[NO], FILE3_10.DESCA, FILE3_10.TITLE, CASE WHEN FILE3_10.FAMILY = 1 THEN 'Ê«·⁄«∆·…' ELSE '' END,CONVERT(VARCHAR(10),MAX(FILE3_10.[DATE_END]),111),FILE3_10.NOTES" & _
          " FROM FILE3_10 LEFT JOIN FILE4_30 ON FILE3_10.CODE = FILE4_30.MEMBER"
If ValidNum(xNo1.text) Then
    cWhere = cWhere & Tr(cWhere, " and ") & "FILE3_10.[NO] " & IIf(ValidNum(xNo2.text), ">=", " = ") & xNo1.text
    If ValidNum(xNo2.text) Then
        aHeader(0) = BetweenString(xNo1.text, xNo2.text, "„‰ —ﬁ„ ﬂ«—‰ÌÂ : ", "Õ Ì —ﬁ„ ﬂ«—‰ÌÂ : ")
    Else
        aHeader(0) = "—ﬁ„ ﬂ«—‰ÌÂ : " & xNo1.text
    End If
End If

If ValidNum(xNo2.text) Then
    cWhere = cWhere & Tr(cWhere) & "FILE3_10.[NO] <= " & xNo2.text
    aHeader(0) = BetweenString(xNo1.text, xNo2.text, "„‰ —ﬁ„ ﬂ«—‰ÌÂ : ", "Õ Ì —ﬁ„ ﬂ«—‰ÌÂ : ")
End If

If ValidNum(xCode.text) Then
    cWhere = cWhere & Tr(cWhere) & "FILE3_10.CODE = " & xCode.text
    aHeader(1) = "«·⁄÷Ê : " & xCodeDesca.Caption
End If

If xSeason.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & "FILE4_30.[YEAR] = " & xSeason.BoundText
    aHeader(1) = "„Ê”„ «·ÿ»«⁄… : " & xSeason.text
End If

If IsDate(xDate1.text) Then
    cWhere = cWhere & Tr(cWhere) & "FILE3_10.DATE_END >= " & DateSq(xDate1.text)
    aHeader(2) = " «—ÌŒ «‰ Â«¡ " & BetweenString(xDate1.text, xDate2.text)
End If
          
If IsDate(xDate2.text) Then
    cWhere = cWhere & Tr(cWhere) & "FILE3_10.DATE_END <= " & DateSq(xDate2.text)
    aHeader(2) = " «—ÌŒ «‰ Â«¡ " & BetweenString(xDate1.text, xDate2.text)
End If

If chkCurrent.Value Then
    cWhere = cWhere & Tr(cWhere) & "(NOT FILE3_10.[NO] IS NULL)"
    aHeader(3) = BetweenString(xDate1.text, xDate2.text)
End If

If Trim(xNotes.text) <> "" Then
    cWhere = cWhere & Tr(cWhere) & MyParnAnd(xNotes.text, "FILE3_10.NOTES")
    aHeader(4) = BetweenString(xDate1.text, xDate2.text)
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " GROUP BY FILE3_10.CODE, FILE3_10.[NO], FILE3_10.DESCA, FILE3_10.TITLE, FILE3_10.FAMILY,FILE3_10.NOTES ORDER BY CASE WHEN FILE3_10.NO IS NULL THEN 1 ELSE 0 END,FILE3_10.NO"
Set data10.Recordset = myCmd(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
'cString = "SELECT FILE3_10.CODE, FILE3_10.[NO], FILE3_10.DESCA, FILE3_10.TITLE, CASE WHEN FILE3_10.FAMILY = 1 THEN 'Ê«·⁄«∆·…' ELSE '' END,MAX(FILE4_30.[DATE]),FILE3_10.NOTES" & _

    .TextMatrix(0, 0) = "„”·”·"
    .TextMatrix(0, 1) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 2) = "—ﬁ„ «·ﬂ«—‰ÌÂ"
    .TextMatrix(0, 3) = "«”„ «·⁄÷Ê"
    .TextMatrix(0, 4) = "«··ﬁ»"
    .TextMatrix(0, 5) = "⁄«∆·Ì"
    .TextMatrix(0, 6) = " «—ÌŒ «‰ Â«¡"
    .TextMatrix(0, 7) = "«·»Ì«‰"
    
    .ColHidden(1) = True
        
    .ColWidth(0) = 800
    .ColWidth(1) = 900
    .ColWidth(2) = 900
    .ColWidth(3) = 2500
    .ColWidth(4) = 2000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1500
    .ColWidth(7) = 2500
    .ColDataType(6) = flexDTDate
    
'    .ColHidden(0) = True

'    .ColDataType(2) = flexDTDouble
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble

    
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    For i = 1 To grid1.rows - 1
        .TextMatrix(i, 0) = i
    Next
    
    SBar1.Panels(1).text = IIf(grid1.rows > 2, "⁄œœ «·”Ã·«  : " & grid1.rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdbankfrm1 = Nothing
End Sub
Private Sub cmdPrintLand_Click()
Dim nRate As Double, i As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 4)
aRow(0) = AddFlag(aRow(0), "text", "«·≈Ã„«·Ì")
Dim nwidth As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = grid1.ColWidth(i) + nwidth
    End If
Next
nRate = 15000 / nwidth
Set PrintGrdNew.myForm = Me
PrintGrdNew.doPrint grid1, nRate, -2, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, True, 11, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'ItemsLookupAll Me, osearchitem, myFlag
End Sub

Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xdesca.text, 1
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberH_LookupAll Me, oSearchMember
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
If Not ValidNum(xCode.text) Then Exit Sub
xCodeDesca.Caption = myField("select DESCA from file3_10 where code = " & addvalue(xCode.text), "DESCA", con) & ""
End Sub
Private Sub xdate1_Validate(Cancel As Boolean)
myValidDate xDate1
End Sub
Private Sub xdate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.rows > 1
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Public Sub myProc()
xCode.text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
oSearchMember.Hide
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xno1_GotFocus()
myGotFocus xNo1
End Sub
Private Sub xno1_LostFocus()
myLostFocus xNo1
End Sub
Private Sub xno2_GotFocus()
myGotFocus xNo2
End Sub
Private Sub xno2_LostFocus()
myLostFocus xNo2
End Sub

