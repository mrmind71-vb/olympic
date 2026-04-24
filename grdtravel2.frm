VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grdTravelfrm2 
   Caption         =   "⁄„·Ì«  ·„ Ì’œ— ·Â« »Ê·Ì’…"
   ClientHeight    =   10905
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   20250
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
   ScaleHeight     =   10905
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   585
      Width           =   5685
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1230
         Picture         =   "grdtravel2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2325
         Picture         =   "grdtravel2.frx":2424
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4545
         Picture         =   "grdtravel2.frx":4C0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdtravel2.frx":7101
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3420
         Picture         =   "grdtravel2.frx":956D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   9450
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   10725
      Begin VB.TextBox xCode 
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
         Height          =   345
         Left            =   8685
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox xCode_sup 
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
         Height          =   345
         Left            =   2430
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox xDate1 
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
         Left            =   8325
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xDate2 
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   6750
         TabIndex        =   3
         Top             =   540
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo xPlace1 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Tag             =   "new"
         Top             =   540
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo xPlace2 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Tag             =   "new"
         Top             =   900
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   945
         Width           =   480
      End
      Begin VB.Label xcode_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Tag             =   "t"
         Top             =   900
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "Õ Ì"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "„‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   525
      End
      Begin VB.Label xCode_sup_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Tag             =   "t"
         Top             =   180
         Width           =   2265
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "«·”«∆ﬁ"
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
         Index           =   3
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   9885
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   270
         Width           =   660
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   10575
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   330
      Left            =   180
      Top             =   135
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
      Left            =   -1845
      Top             =   -135
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -1620
      Top             =   -225
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
      Left            =   -1575
      Top             =   -180
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
      Height          =   7575
      Left            =   45
      TabIndex        =   12
      Top             =   1350
      Width           =   20130
      _cx             =   35507
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   300
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   18
      Top             =   10425
      Visible         =   0   'False
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grdTravelfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFilesave As String
Dim con As New ADODB.Connection
Dim oSearchCar As New Search3, oSearchSup As New Search3, oSearchClient As New Search3
Dim LastSalTable As New ADODB.Recordset
Dim LastImpTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub CmdClear_Click()
DefineText Me
grid1.Rows = 1
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1, Array(1)
End Sub

Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
myload
If grid1.Rows > 1 Then
    grid1.SetFocus
    grid1.Select 1, NextEmpty(grid1, 1, 3, 4)
End If
End Sub
Private Sub cmdPrint_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
Dim aHeader As Variant
cHeader1 = "„ «»⁄… «Ã„«·Ì —Õ·«  Œ·«· › —…"
If IsDate(xdate1.Text) Or IsDate(xDate2.Text) Then aHeader = AddFlag(aHeader, BetweenString(Format(xdate1.Text, "YYYY-MM-DD"), xDate2.Text))
'If IsDate(xDate_Policy1.Text) Or IsDate(xDate_Policy2.Text) Then aHeader = AddFlag(aHeader, BetweenString(xDate_Policy1.Text, xDate_Policy2.Text, "„‰  «—ÌŒ »Ê·Ì’…"))
'If Trim(xCar.Text) <> "" Then aHeader = AddFlag(aHeader, "”Ì«—… —ﬁ„ : " & xCar.Text & turn(xCar_Desca.Caption, " ") & xCar_Desca.Caption & turn(xCar_type.Caption, " ") & xCar_type.Caption)
If Trim(xCode_sup.Text) <> "" Then aHeader = AddFlag(aHeader, "«·„Ê—œ : " & Me.xCode_sup_desca.Caption)
If Trim(xCode.Text) <> "" Then aHeader = AddFlag(aHeader, "«·⁄„Ì· : " & xcode_Desca.Caption)
If Trim(xPlace1.Text) <> "" Then aHeader = AddFlag(aHeader, "„‰ : " & xPlace1.Text)
If Trim(xPlace2.Text) <> "" Then aHeader = AddFlag(aHeader, "Õ Ï : " & xPlace2.Text)
'If Trim(xPolicy.Text) <> "" Then aHeader = AddFlag(aHeader, "—ﬁ„ «·»Ê·Ì’… : " & xPolicy.Text)
If xDriver.MatchedWithList Then aHeader = AddFlag(aHeader, "«·”«∆ﬁ : " & xDriver.Text)
'If Not IsEmpty(aHeader) Then
'    cHeader2 = retHeader(aHeader, 0, 2)
'    cHeader3 = retHeader(aHeader, 2, 2)
'    cHeader4 = retHeader(aHeader, 4, 2)
'End If
'Dim aRow(0) As Variant
'aRow(0) = AddFlag(Empty, "row", 1)
'aRow(0) = AddFlag(aRow(0), "col", 0)
'aRow(0) = AddFlag(aRow(0), "cols", 10)
PrintGrdNew.doprint grid1, 0.84, -3, cHeader1, cHeader2, cHeader3, , False, True, 9, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
openCon con
Set DATA1.Recordset = myRecordSet("Select Code,DescA From driver where driver = 1 order by Desca", con)
Set xDriver.RowSource = DATA1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"
    
Set DATA2.Recordset = myRecordSet("SELECT * FROM PLACE_CODES ORDER BY DESCA", con)
Set xPlace1.RowSource = DATA2
xPlace1.ListField = "DESCA"
xPlace1.BoundColumn = "CODE"
    
Set xPlace2.RowSource = DATA2
xPlace2.ListField = "DESCA"
xPlace2.BoundColumn = "CODE"
        
Set grid1.DataSource = data11
data11.ConnectionString = strCon
Fixgrd
grid1.Rows = 1
LoadText Me
'GetCaption
End Sub
Private Sub myload()
Dim cString As String
cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY,Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),DRIVER.DESCA + CASE WHEN DRIVER_B.DESCA IS NULL THEN '' ELSE '-' + DRIVER_B.DESCA END   ,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL" & _
        " FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE LEFT JOIN FILE6_20 ON TRAVEL_H.DOC_NO = FILE6_20.TRAVEL" & _
        " LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE " & _
        " LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
        " LEFT JOIN DRIVER AS DRIVER_B ON TRAVEL_H.DRIVER2 = DRIVER_B.CODE" & _
        " LEFT JOIN FILE4_10 ON TRAVEL_H.CODE_SUP = FILE4_10.CODE" & _
        " LEFT JOIN TRAVEL_C ON TRAVEL_H.DOC_NO = TRAVEL_C.DOC_NO" & _
        " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
        " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE"

'cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY,Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),DRIVER.DESCA + CASE WHEN DRIVER_B.DESCA IS NULL THEN '' ELSE '-' + DRIVER_B.DESCA END   ,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL,CASE WHEN CODE_SUP IS NULL THEN  SUM(COALESCE(TRAVEL_C.[VALUE],0)) ELSE TOTAL_SUP END,TRAVEL_H.TOTAL - CASE WHEN CODE_SUP IS NULL THEN SUM(COALESCE(TRAVEL_C.[VALUE],0)) ELSE TOTAL_SUP  END,FILE6_20.DOC_NO " & _
'        " FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE LEFT JOIN FILE6_20 ON TRAVEL_H.DOC_NO = FILE6_20.TRAVEL" & _
'        " LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE " & _
'        " LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
'        " LEFT JOIN DRIVER AS DRIVER_B ON TRAVEL_H.DRIVER2 = DRIVER_B.CODE" & _
'        " LEFT JOIN FILE4_10 ON TRAVEL_H.CODE_SUP = FILE4_10.CODE" & _
'        " LEFT JOIN TRAVEL_C ON TRAVEL_H.DOC_NO = TRAVEL_C.DOC_NO" & _
'        " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
'        " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE"
cString = cString & turn(cString) & "(TRAVEL_H.POLICY IS NULL OR TRAVEL_H.DATE_POLICY IS NULL)"
If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE >= " & DateSq(xdate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE <= " & DateSq(xDate2.Text)
End If

If Trim(xCode.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.CODE = " & MyParn(xCode.Text)
End If

If Trim(xCode_sup.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.CODE_SUP = " & MyParn(xCode_sup.Text)
End If


If xPlace1.MatchedWithList Then
    cString = cString & turn(cString) & "PLACE1 = " & xPlace1.BoundText
End If

If xPlace2.MatchedWithList Then
    cString = cString & turn(cString) & "PLACE2 = " & xPlace2.BoundText
End If

If xDriver.MatchedWithList Then
    cString = cString & turn(cString) & "(TRAVEL_H.DRIVER = " & MyParn(xDriver.BoundText) & " OR TRAVEL_H.DRIVER2 = " & MyParn(xDriver.BoundText) & ")"
End If


cString = cString & " GROUP BY  TRAVEL_H.DOC_NO,TRAVEL_H.[DATE],FILE3_10.Desca,TRAVEL_H.POLICY,TRAVEL_H.[DATE_POLICY],DRIVER.DESCA,DRIVER_B.DESCA,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL,TRAVEL_H.CODE_SUP,TRAVEL_H.TOTAL_SUP,FILE6_20.DOC_NO"
Set data11.Recordset = myRecordSet(cString, con)
'Generalarray(1) = Generalarray(1) & " Where (FILE6_20.TRAVEL Is Null)"
'Generalarray(2) = "Order by TRAVEL_H.DATE , TRAVEL_H.DOC_NO "
Fixgrd
End Sub
Sub Fixgrd()
 With grid1
.RowHeight(0) = 700
.WordWrap = True
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 2) = "«· «—ÌŒ"
.TextMatrix(0, 3) = "≈”„ «·⁄„Ì·"
.TextMatrix(0, 4) = "—ﬁ„ «·»Ê·Ì’…"
.TextMatrix(0, 5) = " «—ÌŒ «·»Ê·Ì’…"
.TextMatrix(0, 6) = "≈”„ «·”«∆ﬁ"
.TextMatrix(0, 7) = "—ﬁ„ «·”Ì«—…"
.TextMatrix(0, 8) = "«·„Ê—œ"
.TextMatrix(0, 9) = "„‰"
.TextMatrix(0, 10) = "≈·Ì"
.TextMatrix(0, 11) = "«·‰Ê·Ê‰"

.FrozenCols = 2
.ColWidth(0) = 800
.ColWidth(1) = 900
.ColWidth(2) = 1300
.ColWidth(3) = 2100
.ColWidth(4) = 1400
.ColWidth(5) = 1400
.ColWidth(6) = 2000
.ColWidth(7) = 1500
.ColWidth(8) = 1900
.ColWidth(9) = 1300
.ColWidth(10) = 1300
.ColWidth(11) = 1200
.ColComboList(11) = "..."
.ColDataType(11) = flexDTDouble


.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4

For i = 0 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next

'    .SubtotalPosition = flexSTAbove
'    .Subtotal flexSTSum, -1, 10, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 12, "#0", vbRed, vbYellow, True, "  "
StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 1
'    If .Rows > 1 Then
'        For I = 0 To 9
'            .TextMatrix(1, I) = "«·≈Ã„«·Ì"
'        Next
'        .MergeRow(1) = True
'    End If
'    .MergeCells = flexMergeFree

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Unload Me
Set grditem1 = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
If Not myreplace(Row) Then
    myload
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myload
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If grid1.Row < 1 Then Exit Sub
If grid1.TextMatrix(Row, 9) = "" Or grid1.TextMatrix(Row, 10) = "" Then
    MsgBox "„ﬂ«‰ «·ﬁÌ«„ Ê«·Ê’Ê· €Ì— „”Ã·Ì‰"
    Exit Sub
End If
travel_weightfrm.sDoc_no = grid1.TextMatrix(Row, 1)
travel_weightfrm.Show 1
myload
End Sub

Private Sub Grid1_DblClick()
If grid1.Row < 2 Then Exit Sub
If grid1.Row > 1 Then
    Travelfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 0)
    Travelfrm.Show
End If
End Sub
Sub myProc()
If ActiveControl.Name = xCode_sup.Name Then
    xCode_sup.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    xCode_sup_Validate False
    Unload oSearchSup
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    xCode_Validate False
    Unload oSearchClient
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, grid1.Row, grid1.Col
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 5 Then
    If (Not IsDate(.EditText)) And Trim(.EditText) <> "" Then
        Cancel = True
    Else
        .EditText = Format(.EditText, "YYYY/MM/DD")
    End If
End If
End With
End Sub
Private Sub xCode_sup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub
Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookupAll Me, oSearchClient
End Sub
Private Sub xCode_sup_Validate(Cancel As Boolean)
xCode_sup_desca.Caption = ""
If xCode_sup.Text = "" Then Exit Sub
xCode_sup.Text = RetZero(xCode_sup.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file4_10 where code = " & MyParn(xCode_sup.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xCode_sup_desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Sub xCode_Validate(Cancel As Boolean)
xcode_Desca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file3_10 where code = " & MyParn(xCode.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ«·⁄„Ì· €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xcode_Desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Sub xCode_sup_GotFocus()
myGotFocus xCode_sup
End Sub
Private Sub xCode_sup_LostFocus()
myLostFocus xCode_sup
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDriver_GotFocus()
myGotFocus xDriver
End Sub
Private Sub xDriver_LostFocus()
myLostFocus xDriver
If Not xDriver.MatchedWithList Then xDriver.BoundText = ""
End Sub
Private Sub xPlace1_GotFocus()
myGotFocus xPlace1
End Sub
Private Sub xPlace2_GotFocus()
myGotFocus xPlace2
End Sub
Private Sub xPlace2_LostFocus()
myLostFocus xPlace2
If Not xPlace2.MatchedWithList Then xPlace2.BoundText = ""
End Sub
Private Sub xPlace1_LostFocus()
myLostFocus xPlace1
If Not xPlace1.MatchedWithList Then xPlace1.BoundText = ""
End Sub
Private Sub Grid1_EnterCell()
With grid1
If .Col = 4 Or .Col = 5 Or .Col = 11 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If NextEmpty(grid1, Row, Col) <= 5 Then
    .Col = Col + 1
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 4, 5)
    .ShowCell Row + 1, 1
Else
    .Select Row, Col
End If
End With
End Sub
Private Function myreplace(Row As Long) As Boolean
Dim aInsert As Variant
con.BeginTrans
On Error GoTo myerror
aInsert = AddFlag(Empty, "[POLICY]", Val(grid1.TextMatrix(Row, 4)))
aInsert = AddFlag(aInsert, "[DATE_POLICY]", addDate(grid1.TextMatrix(Row, 5)))
con.Execute addUpdate(aInsert, "TRAVEL_H", "DOC_NO = " & addstring(grid1.TextMatrix(Row, 1)))
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
