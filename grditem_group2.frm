VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grditemGroupfrm 
   Caption         =   " ﬁ—Ì— «·”œ«œ «·ÌÊ„Ì"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleHeight     =   10290
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   6570
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   990
      Width           =   1455
      Begin VB.CheckBox chkGroup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„Ã„Ê⁄« "
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   90
      Width           =   4875
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "„« ⁄œ« »‰Êœ «·›’·"
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
         Index           =   1
         Left            =   1935
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Width           =   1845
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "»‰Êœ «·›’· ›ﬁÿ"
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
         Index           =   2
         Left            =   90
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Width           =   1845
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   0
         Left            =   3960
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   990
      Width           =   7080
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "«Œ—Ì"
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
         Index           =   4
         Left            =   135
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   900
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "›—Êﬁ ﬁÌ„… „÷«›…"
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
         Index           =   3
         Left            =   1395
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Width           =   1665
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "„—ﬂ“ Œœ„« "
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
         Index           =   2
         Left            =   3150
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1305
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "«‘ —«ﬂ«  ›ﬁÿ"
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
         Index           =   1
         Left            =   4545
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1350
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   0
         Left            =   5985
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   810
      Width           =   4920
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2340
         Picture         =   "grditem_group2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grditem_group2.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3555
         Picture         =   "grditem_group2.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "grditem_group2.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   6615
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   8520
      Begin VB.TextBox xdoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox xDoc_No1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Tag             =   "D"
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   3210
         _ExtentX        =   5662
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì’«· —ﬁ„"
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
         Left            =   3435
         TabIndex        =   15
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…"
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
         Index           =   1
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
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
         Index           =   0
         Left            =   7560
         TabIndex        =   10
         Top             =   585
         Width           =   690
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
         Left            =   7575
         TabIndex        =   9
         Top             =   225
         Width           =   660
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   9960
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -180
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   4950
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   7980
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1575
      Width           =   15090
      _cx             =   26617
      _cy             =   14076
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
      Cols            =   5
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -540
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   9765
      Visible         =   0   'False
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grditemGroupfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sDate1 As String, sdate2 As String
Dim aHeader()
Private Sub chkLast_Click()
For i = grid1.rows - 1 To 1 Step -1
    If i < grid1.rows - 10 Then
        grid1.RowHidden(i) = chkLast.Value = 1
    End If
Next
End Sub

Private Sub cmdExel_Click()
Dim sHeader As String, nMargin As Integer

sHeader = Me.Caption
nMargin = 40
If retHeader(aHeader, 0, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 0, 3)
    nMargin = nMargin + 15
End If
If retHeader(aHeader, 1, 3, "-") <> "" Then
    sHeader = sHeader & turn(sHeader, Chr(13)) & retHeader(aHeader, 1, 3, "-")
    nMargin = nMargin + 15
End If


Dim aSplit As Variant
'aSplit = AddFlag(aSplit, "title_col", "A:B")
'aSplit = AddFlag(aSplit, "title_row", "1:1")
'aSplit = AddFlag(aSplit, "center_header", sHeader)
ToFileExel2 grid1, , , , , 1, , , , , , Me

End Sub

Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
Me.MousePointer = 11
If chkGroup.Value = 0 Then
    myload
Else
    myloadGroup
End If
Me.MousePointer = 0
End Sub

Private Sub CmdLast_Click()

End Sub
Private Sub cmdPrint_Click()
Dim i As Long
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", grid1.rows - 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 1)
Set PrintGrdNew.myForm = Me
If chkGroup.Value = 0 Then
    If optType(0) Then
        PrintGrdNew.doPrint Me.grid1, 0.8, -1, Me.Caption, retHeader(aHeader, 0, 2, " - "), retHeader(aHeader, 2, 2), , False, IIf(grid1.Cols > 10, True, False), IIf(grid1.Cols > 10, 9, 11), , aRow, , , 100, 100
    ElseIf optType(1) Then
        PrintGrdNew.doPrint Me.grid1, 0.81, -1, Me.Caption, retHeader(aHeader, 0, 2, " - "), retHeader(aHeader, 2, 2), , False, IIf(grid1.Cols > 10, True, False), IIf(grid1.Cols > 10, 10, 11), , aRow, , , 100, 100
    Else
        PrintGrdNew.doPrint Me.grid1, 1.1, -1, Me.Caption, retHeader(aHeader, 0, 2, " - "), retHeader(aHeader, 2, 2), , False, False, 12, , aRow, , , 100, 100
    End If
Else
    PrintGrdNew.doPrint Me.grid1, 1, -1, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), , False, False, 11, , aRow, , , 100, 100
End If
PrintGrdNew.Show 1
End Sub

Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("SELECT * FROM FILE6_10G ORDER BY DESCA ", con)
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set grid1.DataSource = DATA11
Fixgrd
If IsDate(sDate1) Then xDate1.text = sDate1
If IsDate(sdate2) Then xDate2.text = sdate2
If IsDate(sDate1) Or IsDate(sdate2) Then myload
End Sub
Private Sub myload()
With grid1
        
    Dim cString As String
    Dim cFields As String
    Dim cField As String
    
    cField = "FILE6_20.TOTAL"
'    If optType(0).Value Then
'        cField = "FILE6_20.Total"
'    ElseIf optType(1).Value Then
'        cField = "FILE6_20.Total- FILE6_20.tax"
'    ElseIf optType(2).Value Then
'        cField = "FILE6_20.tax"
'    ElseIf optType(3).Value Then
'        cField = "FILE6_20.Total-FILE6_20.tax"
'    End If

    ReDim aHeader(3)
    cString = "SELECT FILE6_10G.CODE,FILE6_10G.DESCA  " & _
              " FROM  FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_no = FILE6_20.DOC_NO" & _
              " INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM " & _
              " INNER JOIN [PAID_TYPES] ON FILE6_20H.TYPE = PAID_TYPES.CODE" & _
              " LEFT JOIN FILE6_10G ON FILE6_10.[GROUP] = FILE6_10G.CODE"
            
    cString = cString & " WHERE (NOT FILE6_20H.FORM_NO IS NULL)  AND IsFawry = 0 AND " & cField & " <> 0"
    
    If optclose(1).Value Then
        cWhere = cWhere & turn(cWhere, " and ") & "FILE6_20H.DOC_NO IN(SELECT DOC_NO FROM FILE6_20 INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM GROUP BY FILE6_20.DOC_NO HAVING MAX(CAST( FILE6_10.NO_CASH AS INT)) = 0)"
    ElseIf optclose(2).Value Then
        cWhere = cWhere & turn(cWhere, " and ") & " FILE6_10.NO_CASH = 1"
    End If
    
    If optType(0).Value Then
        aHeader(3) = "≈Ã„«·Ì »‰Êœ «‘ —«ﬂ«  Ê„—ﬂ“ Œœ„« "
    ElseIf optType(1).Value Then
        cWhere = cWhere & Tr(cWhere) & " PAID_TYPES.TYPE <> 200 AND PAID_TYPES.TYPE <> 100 AND PAID_TYPES.TYPE <> 300"
        aHeader(3) = "≈Ã„«·Ì »‰Êœ «·«‘ —ﬂ« "
    ElseIf optType(2) Then
        cWhere = cWhere & Tr(cWhere) & " PAID_TYPES.TYPE = 200"
        aHeader(3) = "≈Ã„«·Ì „—ﬂ“ Œœ„« "
    ElseIf optType(3) Then
        cWhere = cWhere & Tr(cWhere) & " PAID_TYPES.TYPE = 100"
        aHeader(3) = "≈Ã„«·Ì ›—Êﬁ ﬁÌ„… „÷«›…"
    ElseIf optType(4) Then
        cWhere = cWhere & Tr(cWhere) & " PAID_TYPES.TYPE = 300"
        aHeader(3) = "„ÿ«·»«  «Œ—Ì"
    End If
    
    If IsDate(xDate1.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE >= " & DateSq(xDate1.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    
    If IsDate(xDate2.text) Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE <= " & DateSq(xDate2.text)
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    End If
    
    If xGroup.MatchedWithList Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_10.[GROUP] = " & xGroup.BoundText
        aHeader(1) = "„Ã„Ê⁄… «·»‰œ : " & xGroup.text
    End If
    
    If Trim(xDoc_No1.text) <> "" Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.FORM_NO >= " & addstring(xDoc_No1.text)
        aHeader(2) = " „‰ „” ‰œ : " & BetweenString(xDoc_No1.text, xdoc_no2.text)
    End If
    
    If Trim(xdoc_no2.text) <> "" Then
        cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.FORM_NO <= " & addstring(xdoc_no2.text)
        aHeader(2) = " „‰ „” ‰œ : " & BetweenString(xDoc_No1.text, xdoc_no2.text)
    End If
    
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    
    
    
    cString = cString & " GROUP BY FILE6_10G.CODE,FILE6_10G.DESCA"
    'cString = cString & " HAVING SUM(" & cField & ") <> 0"
    cString = cString & " ORDER BY FILE6_10G.CODE"
    
    aGroup = GetRowsNew(cString, con)
        
    If Not IsEmpty(aGroup) Then
        For i = 0 To UBound(aGroup)
            If IsNull(retFlag(aGroup(i), "CODE")) Then
                cFields = cFields & turn(cFields, ",") & myiif("FILE6_10.[GROUP] IS NULL", cField) & " AS [»œÊ‰ „Ã„Ê⁄…] "
            Else
                cFields = cFields & turn(cFields, ",") & myiif("FILE6_10.[GROUP] = " & MyParn(retFlag(aGroup(i), "CODE")), cField) & " AS [" & retFlag(aGroup(i), "desca") & "]"
            End If
        Next
    End If
    
    If Not xGroup.MatchedWithList Then
        cFields = cFields & turn(cFields, ",") & "SUM(" & cField & ") AS [«·≈Ã„«·Ì]"
    End If
    
    cString = " SELECT  FILE6_20H.FORM_NO " & turn(cFields, ",") & _
              cFields & _
              " FROM  FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_no = FILE6_20.DOC_NO" & _
              " INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM " & _
              " LEFT JOIN FILE6_10G ON FILE6_10.[GROUP] = FILE6_10G.CODE" & _
              " INNER JOIN [PAID_TYPES] ON FILE6_20H.TYPE = PAID_TYPES.CODE"
    
    cString = cString & " WHERE (NOT FILE6_20H.FORM_NO IS NULL) AND IsFawry = 0"
    If cWhere <> "" Then cString = cString & " AND " & cWhere
    cString = cString & " GROUP BY  FILE6_20H.FORM_NO "
    cString = cString & " HAVING SUM(" & cField & ") <> 0"
    cString = cString & " ORDER BY FILE6_20H.FORM_NO"
    Set DATA11.Recordset = myCmd(cString, con, adText, , 300)
End With
Fixgrd
StatusBar1.Panels(1).text = IIf(grid1.rows - 2 > 0, "⁄œœ «·«” „«—«  : " & grid1.rows - 2, "")
End Sub
Private Sub myloadGroup()
With grid1
Dim cString As String
Dim aPrm As Variant
ReDim aHeader(3)


aPrm = AddFlag(aPrm, "OPTION1", optclose(1).Value)
aPrm = AddFlag(aPrm, "OPTION2", optclose(2).Value)

    
    If IsDate(xDate1.text) Then
        aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    Else
        aPrm = AddFlag(aPrm, "DATE1", Null)
    End If
    
    If IsDate(xDate2.text) Then
        aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
        aHeader(0) = "⁄‰ «·› —… „‰ : " & BetweenString(xDate1.text, xDate2.text)
    Else
        aPrm = AddFlag(aPrm, "DATE2", Null)
    End If
    
    If xGroup.MatchedWithList Then
        aPrm = AddFlag(aPrm, "GROUP", xGroup.BoundText)
        aHeader(1) = "„Ã„Ê⁄… «·»‰œ : " & xGroup.text
    Else
        aPrm = AddFlag(aPrm, "GROUP", Null)
    End If
    
    If Trim(xDoc_No1.text) <> "" Then
        aPrm = AddFlag(aPrm, "FORM_NO1", xDoc_No1.text)
        aHeader(2) = " „‰ „” ‰œ : " & BetweenString(xDoc_No1.text, xdoc_no2.text)
    Else
        aPrm = AddFlag(aPrm, "FORM_NO1", Null)
    End If
    
    If Trim(xDoc_No1.text) <> "" Then
        aPrm = AddFlag(aPrm, "FORM_NO2", xdoc_no2.text)
        aHeader(2) = " „‰ „” ‰œ : " & BetweenString(xDoc_No1.text, xdoc_no2.text)
    Else
        aPrm = AddFlag(aPrm, "FORM_NO2", Null)
    End If
        
    
    Set DATA11.Recordset = myCmd("dbo.sp_items_cash_group", con, adStoredProc, aPrm)
End With
fixgrd2
StatusBar1.Panels(1).text = IIf(grid1.rows - 2 > 0, "⁄œœ «·»‰Êœ : " & grid1.rows - 2, "")
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSort
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 1
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    
    .MergeCells = flexMergeFree
    .ColWidth(0) = 1000
    
    .SubtotalPosition = flexSTBelow
    For i = 1 To grid1.Cols - 1
        .ColWidth(i) = 1190
        .ColDataType(i) = flexDTDouble
        .Subtotal flexSTSum, -1, i, "#0.00", &HC0FFC0, vbBlack, True, "  "
        .ColAlignment(i) = flexAlignCenterTop
    Next
    
    If grid1.rows > 1 Then
        For i = 1 To grid1.Cols - 1
            .TextMatrix(grid1.rows - 1, i) = mRound(.TextMatrix(grid1.rows - 1, i))
        Next
    End If
    .ColWidth(.Cols - 1) = 1300
    
    If grid1.rows > 1 Then
        .TextMatrix(grid1.rows - 1, 0) = "«·«Ã„«·Ì"
    End If
    .Cell(flexcpAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Sub fixgrd2()
    With grid1
    .RowHeight(0) = 300
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ «·»‰œ"
    .TextMatrix(0, 1) = "«”„ «·»‰œ"
    .TextMatrix(0, 2) = "«·«€Ã„«·Ì"
    
    .MergeCells = flexMergeFree
    .ColWidth(0) = 1000
    .ColWidth(1) = 6000
    .ColWidth(2) = 1500
    
    .SubtotalPosition = flexSTBelow
    .ColDataType(2) = flexDTDouble
    .Subtotal flexSTSum, -1, 2, "#0.00", &HC0FFC0, vbBlack, True, "  "
    
    If grid1.rows > 1 Then
        For i = 2 To grid1.Cols - 1
            .TextMatrix(grid1.rows - 1, i) = mRound(.TextMatrix(grid1.rows - 1, i))
        Next
    End If
    
    If grid1.rows > 1 Then
        .TextMatrix(grid1.rows - 1, 0) = "«·«Ã„«·Ì"
    End If
    .Cell(flexcpAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set grditemGroupfrm = Nothing
closeCon con
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub

Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub

Private Sub xDoc_no1_GotFocus()
myGotFocus xDoc_No1
End Sub
Private Sub xDoc_no1_LostFocus()
myLostFocus xDoc_No1
End Sub

Private Sub xDoc_no2_GotFocus()
myGotFocus xdoc_no2
End Sub
Private Sub xDoc_no2_LostFocus()
myLostFocus xdoc_no2
End Sub

Private Sub xgroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
