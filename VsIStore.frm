VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTStore 
   Caption         =   "„ «»⁄… «·«’‰«›"
   ClientHeight    =   10365
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
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄…  «· ﬁ—Ì—"
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo xstore1 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xstore2 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "·Ì” ·Â« —’Ìœ ›Ï"
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
         Index           =   2
         Left            =   2355
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "·Â« —’Ìœ ›Ï"
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
         Index           =   0
         Left            =   2355
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.TextBox xDescItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EAEAEA&
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
      Left            =   3975
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   900
      Width           =   3015
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1275
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1470
      Width           =   1275
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "⁄—÷"
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
      Left            =   2625
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.TextBox xItem 
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
         Left            =   3105
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1350
         Width           =   4590
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   1320
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroupMain 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
         Top             =   960
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·’‰›"
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
         Height          =   195
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4755
         TabIndex        =   15
         Top             =   1425
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…:"
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
         Index           =   1
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1005
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   3
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9555
         TabIndex        =   3
         Top             =   270
         Width           =   360
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Bindings        =   "VsIStore.frx":0000
      Height          =   7560
      Left            =   150
      TabIndex        =   7
      Top             =   2025
      Width           =   14865
      _cx             =   26220
      _cy             =   13335
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   0
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
      Left            =   0
      Top             =   0
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
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "VsTStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BalStore As New ADODB.Recordset
Dim MyStoreTable As New ADODB.Recordset
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Cmd_Print_Click()
    cHead1 = "»Ì«‰ »—’Ìœ «·«’‰«› „Ê“⁄… ⁄·Ï «·„Œ«“‰ "
    cHead2 = " Õ Ï  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY")
    
Dim aHeader(5)
Dim temptable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If IsDate(xDate1.Text) Then
    aHeader(0) = "[" & "Õ Ï  " & xDate1.Text & "]"
End If


If Trim(xSection.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.[SECTION] = " & xSection.BoundText
    aHeader(1) = "«·ﬁ”„" & xGroup.Text & "]"
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.[group] = " & xGroup.BoundText
    aHeader(2) = "„Ã„Ê⁄… " & xGroup.Text & "]"
End If

If Trim(xGroupMain.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_50.[group] = " & xGroupMain.BoundText
    aHeader(3) = "„Ã„Ê⁄… —∆Ì”Ì…" & xGroup.Text & "]"
End If

         

With grid1
    For I = 3 To .Rows - 1
        temptable.AddNew
        temptable!str1 = .TextMatrix(I, 0)
        temptable!str2 = .TextMatrix(I, 1)
        temptable!val2 = Val(.TextMatrix(I, 5))
        For nCol = 6 To .Cols - 1
            temptable.Fields("Str" & 5 + nCol) = .TextMatrix(0, nCol)
            temptable.Fields("VAL" & 5 + nCol) = Val(.TextMatrix(I, nCol))
        Next nCol
        temptable!str21 = " ›’Ì·Ï —’Ìœ «·√’‰«› „Ê“⁄… ⁄·Ï «·„Œ«“‰"
        temptable.Update
        
    Next I
End With

contemp.BeginTrans
contemp.CommitTrans
main.Report1.Destination = crptToWindow
main.Report1.ReportFileName = App.Path & "\Reports\Item_STORE.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1

temptable.Close

Set temptable = Nothing
Set sourcetable = Nothing

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
End Sub
Private Sub Command1_Click()
Dim temptable As ADODB.Recordset
Dim nCol1 As Double
Dim nCol2 As Double
If xStore1.BoundText = "" Or xStore2.BoundText = "" Then Exit Sub

contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
With grid1
    For I = 1 To .Cols - 1
        If .TextMatrix(1, I) = xStore1.BoundText Then nCol1 = I
        If .TextMatrix(1, I) = xStore2.BoundText Then nCol2 = I
    Next I
    For I = 2 To .Rows - 1
        If Val(.TextMatrix(I, nCol1)) > 0 And Val(.TextMatrix(I, nCol2)) = 0 Then
            temptable.AddNew
            temptable!val12 = GetDesca("SELECT [GROUP] FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(I, 0)))
            temptable!str8 = GetDesca("SELECT [DESCA] FROM FILE1_50 WHERE CODE = " & temptable!val12)
            temptable!str1 = .TextMatrix(I, 0)
            temptable!str2 = .TextMatrix(I, 1)
            
            temptable!val2 = Val(GetDesca("SELECT [package] FROM FILE1_10 WHERE item = " & MyParn(.TextMatrix(I, 0))) & "")
            temptable!val3 = .TextMatrix(I, nCol1)
            
            
            temptable!str21 = "»Ì«‰ «’‰«› ·Â« —’Ìœ ›Ï " & xStore1.Text & " Ê ·Ì” ·Â« —’Ìœ ›Ï " & xStore2.Text
            temptable.Update
        End If
    Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.Destination = crptToWindow
    main.Report1.ReportFileName = App.Path & "\Reports\Item_StoreBAL.RPT"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
temptable.Close
Set temptable = Nothing
End Sub

Private Sub Form_Load()
   openCon con
    
    xDate1.Text = Format(Date, "dd-mm-yyyy")
    
    MyStoreTable.Open "SELECT * FROM FILE0_40 ORDER BY CODE ", con, adOpenKeyset, adLockOptimistic, adCmdText
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_50G order by Desca"
    Set xGroupMain.RowSource = DATA2
    xGroupMain.ListField = "Desca"
    xGroupMain.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select Code,DescA From File0_40 order by Desca"
    Set xStore1.RowSource = data5
    xStore1.ListField = "Desca"
    xStore1.BoundColumn = "Code"
    
    Set xStore2.RowSource = data5
    xStore2.ListField = "Desca"
    xStore2.BoundColumn = "Code"
    
    Set grid1.DataSource = data4
    data4.ConnectionString = strCon
        
    FixGrid
         
    grid1.Rows = 2
    grid1.FixedRows = 2
    
End Sub
Private Sub myload()
Dim I As Double
If BalStore.State = adStateOpen Then BalStore.Close

'cStr1 = " TRANSFORM Sum([IN]-[OUT]) AS BALSTORE SELECT FILE1_11.ITEM From FILE1_11  WHERE DATE <= " & datesq(xDate1.Text) & " GROUP BY FILE1_11.ITEM  PIVOT FILE1_11.store "
'BalStore.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText

If IsDate(xDate1.Text) Then cwhere = " date <= " & DateSq(xDate1.Text)
cField1 = myiif(cwhere, "[IN] - [OUT]") & " AS F_BAL"
I = 2

MyStoreTable.MoveFirst
Do Until MyStoreTable.EOF
    cwhere = " STORE = " & MyParn(MyStoreTable!Code)
    cField2 = cField2 & turnFound(cField2, ",") & myiif(cwhere, "[IN] - [OUT]") & " AS Bal" & MyStoreTable!Code
    MyStoreTable.MoveNext
Loop
With grid1
'                           0               1                 2                3               4
    cStrAll = "  select file1_10.item , file1_10.desca , FILE1_10.PRICE3 , FILE1_10.PRICE2 , reorder , " & _
                cField1 & " , " & cField2 & _
    " from ( FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM )  inner join file1_50 on file1_10.[group] = file1_50.code"
    If xGroup.BoundText <> "" Then cStrAll = cStrAll & turn(cString) & "  file1_10.[GROUP]  = " & xGroup.BoundText
    If xGroupMain.BoundText <> "" Then cStrAll = cStrAll & turn(cString) & "  file1_50.[Group]  = " & xGroupMain.BoundText
    If xSection.BoundText <> "" Then cStrAll = cStrAll & turn(cString) & "  [Section] = " & xSection.BoundText
    If xDesca.Text <> "" Then cStrAll = cStrAll & turn(cString) & "  file1_10.DESCA LIKE ('%" & xDesca.Text & "%')   "
    If xItem.Text <> "" Then cStrAll = cStrAll & turn(cString) & "  [FILE1_10.ITEM] = " & MyParn(xItem.Text)
    
    cStrAll = cStrAll & " GROUP BY FILE1_10.ITEM , FILE1_10.DESCA , FILE1_10.PRICE2 , FILE1_10.PRICE3 , FILE1_10.reorder  "
    data4.RecordSource = cStrAll
    data4.Refresh
End With
FixGrid
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid()
With grid1
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·’‰›"
    .TextMatrix(0, 2) = "”⁄— Ã„·…"
    .TextMatrix(0, 3) = "”⁄— „” Â·ﬂ"
    .TextMatrix(0, 4) = "Õœ «·ÿ·»"
    .TextMatrix(0, 5) = "«·—’Ìœ "
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColHidden(2) = True
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    MyStoreTable.MoveFirst
    I = 6
    Do Until MyStoreTable.EOF
        If I >= grid1.Cols Then grid1.Cols = grid1.Cols + 1
        .TextMatrix(0, I) = MyStoreTable!Desca
        .TextMatrix(1, I) = MyStoreTable!Code
        MyStoreTable.MoveNext
        I = I + 1
    Loop


'    For I = 1 To BalStore.Fields.Count - 1
'        .Cols = .Cols + 1
'        .TextMatrix(1, .Cols - 1) = BalStore.Fields(I).Name
'        MyStoreTable.Find " CODE = " & MyParn(BalStore.Fields(I).Name), , adSearchForward, adBookmarkFirst
'        If Not MyStoreTable.EOF Then .TextMatrix(0, .Cols - 1) = MyStoreTable!DESCA
'    Next I
'
'    For I = 2 To .Rows - 1
'        BalStore.Find " ITEM = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
'        If Not BalStore.EOF Then
'            For nCol = 1 To BalStore.Fields.Count - 1
'                .TextMatrix(I, 5 + nCol) = BalStore.Fields(.TextMatrix(1, 5 + nCol)) & ""
'            Next nCol
'        End If
'    Next I
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    For I = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
            
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    MyStoreTable.Close
    BalStore.Close
    Set MyStoreTable = Nothing
    Set BalStore = Nothing
    closeCon con
End Sub
Private Sub grid1_DblClick()
        Load StoreMove
        StoreMove.xItem.Text = grid1.TextMatrix(grid1.Row, 0)
        StoreMove.xStore.BoundText = grid1.TextMatrix(1, grid1.col)
        StoreMove.Show
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookupAll Me, Search3
End Sub
Private Sub xitem_LostFocus()
xDescItem.Text = GetDesca("select DESCA from file1_10 where item = " & MyParn(xItem.Text)) & ""
End Sub
Sub myProc()
   xItem.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
   xDescItem.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
   Search3.Hide
End Sub

