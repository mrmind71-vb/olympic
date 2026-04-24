VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form trustfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÊÌ«  «·—Õ·« "
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   19725
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAddTravel 
      Caption         =   "«÷«›… «·⁄„·Ì« "
      Height          =   465
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1530
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   12375
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   -45
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "trust.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
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
         Height          =   510
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "trust.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   7740
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "trust.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   12
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "trust.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   13
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "trust.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   14
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
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
         Picture         =   "trust.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "trust.frx":1CE87
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -315
      Top             =   -135
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   -4680
      Top             =   90
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
      Left            =   -3285
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -1260
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   1170
      Top             =   -135
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -5175
      Top             =   -135
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   11250
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   8340
      Begin VB.TextBox xTrust_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2925
         MaxLength       =   300
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   900
         Width           =   4200
      End
      Begin VB.TextBox xBox 
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
         Left            =   6030
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   345
         Left            =   6030
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox xDate 
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
         Height          =   345
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ”ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   900
         Width           =   870
      End
      Begin VB.Label xBox_desca 
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   3075
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   885
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰…"
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
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   -3285
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
   Begin VB.Frame Frame4 
      Caption         =   "„” ‰œ«  «· ”ÊÌ…"
      Height          =   5730
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2025
      Width           =   16710
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5280
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   16530
         _cx             =   29157
         _cy             =   9313
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
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
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   7695
      Width           =   11265
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label xTotal_Desca 
         Caption         =   "«·»«ﬁÌ"
         Height          =   240
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   270
         Width           =   600
      End
      Begin VB.Label xTotal_Cost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label xTotal_trust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8370
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "≈Ã„«·Ì «·⁄Âœ…"
         Height          =   285
         Left            =   9990
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "≈Ã„«·Ì «·„’—Ê›"
         Height          =   285
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   1365
      End
   End
End
Attribute VB_Name = "trustfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, bSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String
Dim cFilter As String, cList As String
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchBox As New Search3, oSearch_Travel As New Search3, oSearchSup As New Search3, oSearchDriver As New Search3
'Dim oAdd As New trust_addfrm
Dim bedit As Boolean
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Private Function myReplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "BOX", addstring(xBox.Text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "[TRUST_NO]", addstring(xTrust_No.Text))
On Error GoTo myerror
con.BeginTrans
If xDoc_No.Text = "" Then
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(sUserName))
    xDoc_No.Text = RetZero(Val(Newflag("TRUST_H", "doc_no")))
    aInsert = AddFlag(aInsert, "[DOC_NO]", addstring(xDoc_No.Text))
    con.Execute addInsert(aInsert, "TRUST_H")
Else
    con.Execute addUpdate(aInsert, "TRUST_H", "doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd Row
con.CommitTrans
myReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0), , 1)
    If nFound <> -1 Then
         MsgBox "»Ê·Ì’… «·‘Õ‰ „ÊÃÊœ… ›Ï «·”ÿ— " & nFound
         Exit Sub
    End If
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    grid1.TextMatrix(grid1.Row, 1) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 3)
    grid1.TextMatrix(grid1.Row, 3) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 4)
    grid1.TextMatrix(grid1.Row, 4) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 6)
    grid1.TextMatrix(grid1.Row, 5) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 7)
    grid1.TextMatrix(grid1.Row, 6) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 8)
    grid1.TextMatrix(grid1.Row, 7) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 9)
    grid1.TextMatrix(grid1.Row, 8) = oSearch_Travel.grid1.TextMatrix(oSearch_Travel.grid1.Row, 10)
    Grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        Unload oSearch_Travel
        CellPos 13, grid1.Row, 2
    Else
        grid1.Select grid1.Rows - 1, 2
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    Unload oSearchDoc
ElseIf ActiveControl.Name = xBox.Name Then
    xBox.Text = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 0)
    xBox_desca.Caption = oSearchBox.grid1.TextMatrix(oSearchBox.grid1.Row, 1)
    Unload oSearchBox
    SendKeys "{TAB}"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdCash_Click()
charge_Cashfrm.Show 1
End Sub
Private Sub cmdCost_Click()
cost_fixfrm.Show 1
End Sub
Private Sub cmdDaySales_Click()
casher_closefrm.Show 1
End Sub
Private Sub cmdCargo_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sCode = xCargo.BoundText
oFlagfrm.sTable = "CARGO_CODES"
oFlagfrm.sCaption = "«‰Ê«⁄ «·Õ„Ê·« "
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
data3.Refresh
xCargo.BoundText = sCode
If Not xCargo.MatchedWithList Then xCargo.BoundText = ""
End Sub
Private Sub cmdAddTravel_Click()
If Trim(xBox.Text) = "" Then Exit Sub
Set oAdd.myForm = Me
oAdd.Sbox = xBox.Text
oAdd.Show 1
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From TRUST_DOC where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRUST_CASH where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From TRUST_H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
       CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
       If CardTable.BOF Then CardTable.MoveFirst
       myLoad
    End If
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myLoad
End Sub

Private Sub cmdGroup_Click()

End Sub

Private Sub CmdInform_Click()
Trust_LookupAll Me, oSearchDoc, cFilter
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myLoad
End If
End Sub
Private Sub CmdAdd_Click()
bAddnew = True
mydefine
On Error Resume Next
Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
mysave
On Error Resume Next
Err.Clear
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Command2_Click()
TDaySal.Show 1
End Sub
Private Sub Command3_Click()
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset
loctable.Open "Select * FROM TRAVEL_H", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If Not IsNull(loctable!Time) Then
        dDate = Format(IIf(Val(Format(loctable!Time, "hh")) > 4, loctable!Time, DateAdd("d", -1, loctable!Time)), "YYYY-MM-DD")
        cString = "update TRAVEL_H set TRAVEL_H.date = " & DateSq(dDate)
        cString = cString & turn(cString) & " doc_no = " & MyParn(loctable!DOC_NO)
        con.Execute cString
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing


Set loctable = New ADODB.Recordset
loctable.Open "Select DOC_NO,SUM(PRICE * QUANT) AS TOTAL FROM TRAVEL GROUP BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    con.Execute "UPDATE TRAVEL_H SET TRAVEL_H.CASH = " & Val(loctable!TOTAL & "") & " WHERE DOC_NO = " & MyParn(loctable!DOC_NO)
    loctable.MoveNext
Loop
MsgBox "done..."
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()

End Sub
Private Sub Form_Activate()
On Error Resume Next
If xDoc_No.Tag = LoadMode Then grid1.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    Grid1_Validate False
    cmdSave_Click
    KeyCode = 0
ElseIf KeyCode = 115 Then
    itemsgrdfrm.Show 1
End If
End Sub
Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con
cList = StrList("SELECT * FROM PLACE_CODES ORDER BY DESCA")
Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing

closeCon con

Unload oSearch
Unload oSearchDoc
Unload oSearchBox
Set trustfrm = Nothing
Err.Clear
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Trim(xBox.Text) = "" Then Exit Sub
Travel_Trust_LookupAll Me, oSearch_Travel, xBox.Text
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bedit Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "DELETE FROM TRUST_DOC WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        myRemove grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
CalcTotals
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then myAddItem

If myReplace(Row) Then
    HandleCntEdit
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
        myloadgrd
    End If
Else
    myloadgrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 2) = ""
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Sub
If Not IsEmpty(aRet) Then
    
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 And Col <> 6 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
End Sub
Private Sub Grid1_EnterCell()
With grid1
If .Col = 1 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
'If grid1.Rows < 2 Then Exit Sub
'If grid1.Row = 0 Then
'    grid1.Row = 1
'    grid1.Col = 1
'End If
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    myRemove grid1.Row
End If
End Sub
Private Sub fixGrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "»Ê·Ì’… «·‘Õ‰|" & "«· «—ÌŒ|" & "„‰|" & "≈·Ì|" & "«·„’—Ê›|" & "«·⁄Âœ…|" & "«·»«ﬁÌ|"
.ColWidth(0) = 500
.ColWidth(1) = 1000
.ColWidth(2) = 1200
.ColWidth(3) = 1400
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColWidth(6) = 1000
.ColWidth(7) = 1050
.ColWidth(8) = 1050
.ColWidth(9) = 1050
.ColComboList(1) = "..."
.ColComboList(4) = cList
.ColComboList(5) = cList
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 8 Then
    .Col = Col + 1
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, 1
    .ShowCell Row + 1, 1
Else
    .Select Row, Col
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
MakeSerial
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_No.Text))
        aInsert = AddFlag(aInsert, "TRAVEL", addstring(grid1.TextMatrix(i, 1)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "TRUST_DOC")
        Else
            con.Execute addUpdate(aInsert, "TRUST_DOC", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
cString = "SELECT TRUST_DOC.TRAVEL,TRAVEL_H.POLICY,CONVERT(VARCHAR(10),TRAVEL_H.DATE_POLICY,111),TRAVEL_H.PLACE1,TRAVEL_H.PLACE2,TRAVEL_BAL.CHARGE,TRAVEL_BAL.TRUST,-1 * TRAVEL_BAL.BALANCE,TRUST_DOC.ID " & _
          " FROM TRUST_DOC INNER JOIN TRAVEL_H ON TRUST_DOC.TRAVEL = TRAVEL_H.DOC_NO LEFT JOIN TRAVEL_BAL ON (TRUST_DOC.TRAVEL = TRAVEL_BAL.DOC_NO AND TRAVEL_BAL.BOX = " & MyParn(xBox.Text) & ")"
cString = cString & turn(cString) & " TRUST_DOC.DOC_NO = " & MyParn(xDoc_No.Text)
DATA11.RecordSource = cString
DATA11.Refresh
myAddItem
End With
CalcTotals
fixGrd
End Sub
Private Sub xBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then BoxLookupAll Me, oSearchBox
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xCode_sup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub
Private Sub xbox_Validate(Cancel As Boolean)
xBox_desca.Caption = ""
If xBox.Text = "" Then Exit Sub
xBox.Text = RetZero(xBox.Text, 6)
Dim aRet As Variant
aRet = GetFields("select code,desca from file0_50 where code = " & MyParn(xBox.Text))
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·Œ“‰… €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xBox_desca.Caption = retFlag(aRet, "desca") & ""
End If
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
Dim i As Integer
If Not bedit Then Exit Function
If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Trim(xBox.Text) = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If
With grid1
End With
MYVALID = True
End Function
Private Sub myLoad()
xDoc_No.Text = CardTable!DOC_NO
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
xBox.Text = CardTable!BOX & ""
xTrust_No.Text = CardTable!trust_no & ""
xBox_desca.Caption = CardTable!Box_Desca & ""
myloadgrd
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Handlecontrols LoadMode
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Sub mydefine()
xDoc_No.Text = ""
xTrust_No.Text = ""
xDate.Text = ""
xBox.Text = ""
xBox_desca.Caption = ""
xTotal_Cost.Caption = ""
xTotal_trust.Caption = ""
xRest.Caption = ""
grid1.Rows = 1
myAddItem
fixGrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdAdd.Enabled = nMode = LoadMode And bedit
cmdSave.Enabled = (bedit)
CmdDel.Enabled = nMode = LoadMode And bedit
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
CmdInform.Enabled = sDoc_no = ""
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
xBox.Enabled = Not (grid1.Rows > 2)
End Sub

Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub

Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub

Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myLoad
ElseIf xDoc_No.Tag = LoadMode Then
    mydefine
End If
End Sub
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotal_Cost As Double, nTotal_Trust As Single, nRest As Double
With grid1
For i = 1 To .Rows - 2
    nTotal_Trust = Val(.TextMatrix(i, 6)) + nTotal_Trust
    nTotal_Cost = Val(.TextMatrix(i, 7)) + nTotal_Cost
    nRest = Val(.TextMatrix(i, 8)) + nRest
Next
xTotal_trust.Caption = Myvalue(nTotal_Trust)
xTotal_Cost.Caption = Myvalue(nTotal_Cost)
xRest.Caption = Myvalue(nRest)
End With
End Function
Private Function mysave() As Boolean
If Not MYVALID Then Exit Function
CalcTotals
If Not myReplace Then Exit Function
Inform " „ Õ›Ÿ «·„” ‰œ"
If sDoc_no <> "" Then
    Unload Me
    Exit Function
End If
openCardTable
myUndo
End Function
Private Function doprint() As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!Date1 = DateFix(xDate.Text)
    temptable!str5 = TurnValue(xtime.Caption)
    
    If xNotes.Text <> "" Then temptable!str2 = "«·”«œ… : " & xNotes.Text
    temptable!str11 = "iPlanet "
    
    
    temptable!str3 = ArbString(Val(xDoc_No.Text))
    temptable!str6 = .TextMatrix(i, 2)
    temptable!str4 = TurnValue(xBox.Text)
    temptable!val1 = Val(.TextMatrix(i, 3))
    temptable!val2 = Val(.TextMatrix(i, 4))
    temptable!Val3 = Val(.TextMatrix(i, 5))
    temptable!str14 = TurnValue(cComp_Name)
    temptable!str15 = TurnValue(cComp_address)
    temptable!str16 = TurnValue(turn(cComp_Phone, "Phone : ") & cComp_Phone)

    temptable!val4 = Val(xTotal_Item.Caption)
    temptable!Val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xCash.Caption)
    temptable!VAL7 = Val(xPay.Caption)
    temptable!VAL8 = Val(xRest.Caption)
    temptable.Update
Next i
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.Destination = crptToPrinter
main.REPORT1.ReportFileName = App.Path & "\Reports\sales_bon.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
main.REPORT1.Destination = crptToWindow
doprint = True
GoTo closeCon
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon:
temptable.Close
Set temptable = Nothing
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
cmdSave.Enabled = (bedit)
xBox.Enabled = Not (grid1.Rows > 2)
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TRUST_H.*,FILE0_50.DESCA AS BOX_DESCA FROM TRUST_H INNER JOIN FILE0_50 ON TRUST_H.BOX = FILE0_50.CODE"
If sDoc_no <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myLoad
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xtotal_GotFocus()
myGotFocus xTotal
End Sub
Private Sub xtotal_LostFocus()
myLostFocus xTotal
CalcTotals
End Sub
Private Sub xGas_GotFocus()
myGotFocus xGas
End Sub
Private Sub xGas_LostFocus()
myLostFocus xGas
'Calctotals2
End Sub
Private Sub xCar_GotFocus()
myGotFocus xcar
End Sub
Private Sub xCar_LostFocus()
myLostFocus xcar

End Sub
Private Sub xDistance_GotFocus()
myGotFocus xDistance
End Sub
Private Sub xDistance_LostFocus()
myLostFocus xDistance
'Calctotals2
End Sub
Private Sub xPlace2_GotFocus()
myGotFocus xPlace2
End Sub
Private Sub xPlace2_LostFocus()
myLostFocus xPlace2
End Sub
Private Sub xPlace1_GotFocus()
myGotFocus xPlace1
End Sub
Private Sub xPlace1_LostFocus()
myLostFocus xPlace1
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
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
Private Sub xDriver_GotFocus()
myGotFocus xDriver
End Sub
Private Sub xDriver2_GotFocus()
myGotFocus xDriver2
End Sub
Private Sub xDriver2_LostFocus()
myLostFocus xDriver2
End Sub
Private Sub Xweight_GotFocus()
myGotFocus xWeight
End Sub
Private Sub Xweight_LostFocus()
myLostFocus xWeight
CalcTotals
End Sub
Private Sub xFollower_GotFocus()
myGotFocus xFollower
End Sub
Private Sub xFollower_LostFocus()
myLostFocus xFollower
If Not xFollower.MatchedWithList Then xFollower.BoundText = ""
End Sub
Private Sub xCargo_GotFocus()
myGotFocus xCargo
End Sub
Private Sub xCargo_LostFocus()
myLostFocus xCargo
If Not xCargo.MatchedWithList Then xCargo.BoundText = ""
End Sub
Private Sub xTotal_sup_GotFocus()
myGotFocus xTotal_sup
End Sub
Private Sub xTotal_sup_LostFocus()
myLostFocus xTotal_sup
CalcTotals
End Sub
Private Sub xCode_sup_GotFocus()
myGotFocus xCode_sup
End Sub
Private Sub xCode_sup_LostFocus()
myLostFocus xCode_sup
CalcTotals
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
MakeSerial
Handlecontrols xDoc_No.Tag
CalcTotals
End Sub
Sub Addproc()
With oAdd.grid1
For i = 1 To .Rows - 1
    If Val(.TextMatrix(i, .Cols - 1)) <> 0 Then
        grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
        grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(i, 0)
        grid1.TextMatrix(grid1.Rows - 1, 2) = .TextMatrix(i, 3)
        grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(i, 4)
        grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(i, 5)
        grid1.TextMatrix(grid1.Rows - 1, 5) = .TextMatrix(i, 6)
        grid1.TextMatrix(grid1.Rows - 1, 6) = .TextMatrix(i, 7)
        grid1.TextMatrix(grid1.Rows - 1, 7) = .TextMatrix(i, 8)
        grid1.TextMatrix(grid1.Rows - 1, 8) = .TextMatrix(i, 9)
        grid1.AddItem ""
    End If
Next
End With
Unload oAdd
mysave
'On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub

