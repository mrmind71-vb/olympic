VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Addressesfrm 
   Caption         =   " ·Ì›Ê‰«  «·«⁄÷«¡"
   ClientHeight    =   9975
   ClientLeft      =   -2415
   ClientTop       =   2415
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1485
      Width           =   3120
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   315
      Width           =   5550
      Begin VB.CommandButton cmdChange 
         Caption         =   "  ÕÊÌ· «·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo xRegion2 
         Height          =   330
         Left            =   1350
         TabIndex        =   17
         Top             =   225
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label2 
         Caption         =   "«· ﬁ”Ì„ «·«œ«—Ì"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1035
      Width           =   5550
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3330
         Picture         =   "addresses.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "addresses.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4410
         Picture         =   "addresses.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2250
         Picture         =   "addresses.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1140
         Picture         =   "addresses.frx":9573
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9600
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
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
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   12555
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   7620
      Begin VB.TextBox xAddress2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1305
         Width           =   5370
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   270
         Width           =   510
      End
      Begin VB.CheckBox xNoRegion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰  ﬁ”Ì„ «œ«—Ì"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   990
         Width           =   2355
      End
      Begin VB.TextBox xPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   6000
      End
      Begin VB.TextBox xAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   5370
      End
      Begin MSDataListLib.DataCombo xRegion 
         Height          =   330
         Left            =   2925
         TabIndex        =   2
         Top             =   945
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label4 
         Caption         =   "«·⁄‰Ê«‰ »œÊ‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label27 
         Caption         =   "«· ﬁ”Ì„ «·«œ«—Ì"
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "«· ·Ì›Ê‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "«·⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   -1125
      Top             =   630
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   -1215
      Top             =   900
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
      Height          =   375
      Left            =   -990
      Top             =   810
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   375
      Left            =   -1350
      Top             =   810
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   420
      Left            =   -1845
      Top             =   720
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   420
      Left            =   -1260
      Top             =   405
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   13
      Top             =   0
      Width           =   0
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7215
      Left            =   90
      TabIndex        =   3
      Top             =   1800
      Width           =   20085
      _cx             =   35428
      _cy             =   12726
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
      Cols            =   6
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   9270
      Visible         =   0   'False
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   582
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "Addressesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileSave As String, cFilePrint As String
Dim con As New ADODB.Connection
Dim oSearch As New Search
Dim aHeader()
Dim printTable As New ADODB.Recordset
Private Sub CmdDel_Click()
End Sub

Private Sub cmdChange_Click()
If MsgBox(" ⁄œÌ· »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
prog1.Value = 0
prog1.Visible = True
For I = 1 To grid1.rows - 1
    prog1.Value = mRound((I / (grid1.rows - 1)) * 100, 2)
    con.Execute "update file1_10 set file1_10.region = " & addvalue(xRegion2.BoundText) & " where file1_10.code = " & addvalue(grid1.TextMatrix(I, 0))
Next
prog1.Value = 0
prog1.Visible = False
myloadgrd
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
End Sub

Private Sub cmdExel_Click()
grid1.ColHidden(3) = True
ToFileExel grid1, , , , , 0.9, , , , 12, , Me
grid1.ColHidden(3) = False
End Sub

Private Sub cmdGo_Click()
myloadgrd
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
Unload Me
End Sub
Private Sub cmdClear_Click()
DefineText Me
grid1.rows = 1
End Sub
Private Sub cmdPrint_Click()
Set PrintGrdNew.myForm = Me
PrintGrdNew.doprint grid1, 0.74, -3, Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 2), False, False, 8, , aRow
grid1.ColHidden(0) = False
PrintGrdNew.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & MainPath & "\MDB\Data.mdb"
openCon con

Set data1.Recordset = myRecordSet(addSelect & "select code,desca from Region_Codes order by desca", con)
Set xRegion.RowSource = data1
xRegion.ListField = "Desca"
xRegion.BoundColumn = "Code"

Set xRegion2.RowSource = data1
xRegion2.ListField = "Desca"
xRegion2.BoundColumn = "Code"


Set grid1.DataSource = data10
Fixgrd
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set grdPhonesfrm = Nothing
'SaveText Me
End Sub
Private Function CountGrid() As Integer
With grid1
For I = 1 To grid1.rows - 1
    CountGrid = CountGrid + 1
Next
End With
End Function
Private Sub countPrint()
nCountPrint = 0
With grid1
For I = 1 To .rows - 1
   If .TextMatrix(I, 6) = True Then nCountPrint = nCountPrint + 1
Next
lblCount.Caption = nCountPrint / 1
End With
End Sub
Sub myProc()
ActiveControl.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Private Sub myloadgrd()
Dim nRecordcount As Long
ReDim aHeader(6)
Dim cString As String

cString = "select file1_10.Code,file1_10.Desca,region_codes.desca,file1_10.Address,file1_10.Phone,file1_10.Mobil from file1_10 left join region_codes on file1_10.region = region_codes.code"
If xPhone.text <> "" Then cWhere = "FILE1_10.PHONE LIKE " & MyParn(xPhone.text & "%")


If Check2.Value = 1 Then
    'cWhere = cWhere & turn(cWhere, " and ") & "(not (isNull(mobil))))"
    cOr = cOr & turn(cOr, " and ") & "len(LTRIM(Rtrim(mobil))) = 11"
    cOr = cOr & turn(cOr, " and ") & "left(LTRIM(Rtrim(mobil)),2) = '01'"
    cOr = "(" & cOr & ")"
    
    
    'cor2 = cor2 & turn(cor2, " and ") & "len(LTRIM(Rtrim(phone))) = 11"
    'cor2 = cor2 & turn(cor2, " and ") & "left(LTRIM(Rtrim(phone)),2) = '01'"
    'cOr2 = "Not Phone is null"
    'cOr2 = "(" & cOr2 & ")"

    'cWhere = cWhere & turn(cWhere, " and ") & "(" & cOr & " Or " & cOr2 & ")"
    cWhere = cWhere & turn(cWhere, " and ") & cOr

End If

If xAddress.text <> "" Then
    If Check1.Value = 0 Then
        cWhere = cWhere & turn(cWhere, " and ") & MyParnAnd(xAddress.text, "FILE1_10.ADDRESS")
    Else
        cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.ADDRESS LIKE " & MyParn("%" & xAddress.text & "%")
    End If
End If
If xAddress2.text <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & MyParnAndNo(xAddress2.text, "FILE1_10.ADDRESS")
End If
If (xRegion.MatchedWithList) And xRegion.BoundText <> "" Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.REGION = " & addvalue(xRegion.BoundText)
If xNoRegion = 1 Then cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.REGION IS NULL"
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " ORDER BY FILE1_10.DESCA"

Set data10.Recordset = myRecordSet(cString, con)
Fixgrd
prog1.Visible = False
Me.MousePointer = 0
CalcTotals
End Sub
Private Sub Fixgrd()
With grid1
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 1) = "«”„ «·⁄÷Ê"
    .TextMatrix(0, 2) = "«· ﬁ”Ì„ «·«œ«—Ì"
    .TextMatrix(0, 3) = "«·⁄‰Ê«‰"
    .TextMatrix(0, 4) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 5) = "«·„Õ„Ê·"
    
    grid1.ColDataType(0) = flexDTLong
    grid1.ColSort(0) = flexSortGenericAscending
      
    .ColWidth(0) = 1000
    .ColWidth(1) = 3500
    .ColWidth(2) = 1500
    .ColWidth(3) = 8000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    For I = 0 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub CalcTotals()
Dim nAll As Long, nPhoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(1).text = ""
If grid1.rows = 1 Then Exit Sub
StatusBar1.Panels(1).text = "⁄œœ «·”Ã·«  : " & grid1.rows - 1
End Sub


Private Sub Grid1_Keyup(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    grid1.RemoveItem grid1.Row
End If
End Sub
