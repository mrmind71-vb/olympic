VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cardqry_hfrm 
   Caption         =   "ÿ»«⁄… ﬂ«—‰ÌÂ«  «·⁄÷ÊÌ… «·‘—›Ì…"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18660
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   18660
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·€«¡ «·œ⁄Ê…"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1935
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   1710
      Width           =   1500
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   13815
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   -45
      Width           =   4605
      Begin VB.CommandButton cmdLastFillgrd 
         Caption         =   "«” —Ã«⁄ «Œ— ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton cmdSavePrint 
         Caption         =   " „  «·ÿ»«⁄…"
         Height          =   390
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry_h.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry_h.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8145
      Width           =   7665
      Begin ComctlLib.ProgressBar prog1 
         Height          =   510
         Left            =   45
         TabIndex        =   36
         Top             =   180
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   900
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13635
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8145
      Width           =   2310
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "«·’›"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "«·⁄„Êœ"
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
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "÷»ÿ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15975
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   8145
      Width           =   2490
      Begin VB.TextBox xRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   315
         Width           =   435
      End
      Begin VB.TextBox xDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label8 
         Caption         =   "«”›·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label9 
         Caption         =   "Ì„Ì‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10575
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8145
      Width           =   3030
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " „  ÿ»«⁄ Â"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   315
         Width           =   1005
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   0
         Left            =   1170
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»œÊ‰ ’Ê—…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   315
         Width           =   915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   1
         Left            =   2655
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.Frame Frame5 
      Height          =   780
      Left            =   5445
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1305
      Width           =   6855
      Begin VB.CommandButton cmdAdd 
         Height          =   555
         Left            =   5175
         Picture         =   "cardqry_h.frx":4D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1680
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "cardqry_h.frx":71F8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   3465
         Picture         =   "cardqry_h.frx":99E3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
      End
      Begin Threed.SSCommand cmdprintrep 
         Height          =   555
         Left            =   1605
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cardqry_h.frx":BE0D
         Caption         =   "ÿ»«⁄…  ﬁ—Ì—"
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "„Ê”„ «·ÿ»«⁄…"
      Height          =   780
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1305
      Width           =   1905
      Begin VB.TextBox xSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·–Ì ·„ Ìÿ»⁄ ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9990
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   990
      Width           =   2265
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   9855
      Width           =   18660
      _ExtentX        =   32914
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   1170
      Top             =   360
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   1845
      Top             =   -45
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
      Left            =   135
      Top             =   135
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
      Left            =   2430
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   420
      Left            =   0
      Top             =   135
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
      Left            =   0
      Top             =   135
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   12330
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   630
      Width           =   6135
      Begin VB.TextBox xNo2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox xNo1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox xCode1 
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
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   945
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "„‰ —ﬁ„"
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   690
      End
      Begin VB.Label xcode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   3840
      End
      Begin VB.Label Label10 
         Caption         =   "«·Ì —ﬁ„"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   690
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "cardqry_h.frx":E411
      Height          =   6000
      Left            =   135
      TabIndex        =   39
      Top             =   2115
      Width           =   18330
      _cx             =   32332
      _cy             =   10583
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
      BackColorFixed  =   16761024
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
      Cols            =   8
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
   Begin VB.Frame Frame8 
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   8100
      Width           =   2760
      Begin VB.CommandButton Command1 
         Caption         =   "«Œ Ì«— «·ﬂ·"
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   2670
      End
   End
End
Attribute VB_Name = "cardqry_hfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileSave As String, cFileSave2 As String, cFilePrint As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim printTable As New ADODB.Recordset
Private Sub cmdAdd_Click()
'checkErr
myloadgrd
'On Error Resume Next
grid1.SaveGrid cFileSave, flexFileData
Err.Clear
cmdPrint.Enabled = (grid1.rows > 1)
checkPhoto
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·ﬂ· !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
    grid1.rows = 1
    grid1.SaveGrid cFileSave, flexFileData
'    DefineText Me
    CalcTotals
End If
End Sub

Private Sub cmdExel_Click()
For I = 1 To grid1.rows - 1
    If Not validPhoto(RetPhotoh(grid1.TextMatrix(I, 2))) Then
        grid1.RowHidden(I) = True
    End If
Next
ToFileExel grid1, , , , , 0.9
For I = 1 To grid1.rows - 1
    grid1.RowHidden(I) = False
Next
End Sub

Private Sub cmdGo_Click()

End Sub

Private Sub cmdPrint_Click()
If grid1.rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
    Exit Sub
End If

If Val(xRow.text) > 5 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.text) > 2 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If

If Not doprint Then
    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
    Exit Sub
End If

'Set cardPrint_one.myForm = Me
'cardPrint_one.PrintArray
'cardPrint_one.Show 1

Set CardPrintNew.myForm = Me
CardPrintNew.PrintArray
CardPrintNew.Show 1

grid1.SaveGrid cFileSave2, flexFileData
If MsgBox(" „  «·ÿ»«⁄… ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    SavePrinted
    For I = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0 Then
            grid1.RemoveItem I
        End If
    Next
    checkPhoto
End If
'grid1.SaveGrid cFilesave, flexFileData
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdClear_Click()
grid1.rows = 1
End Sub
Private Sub cmdMember_Click()
MEMBER.Show 1
End Sub
Private Sub CmdLastFillGrd_Click()
Dim fs As New FileSystemObject
If fs.FileExists(cFileSave) Then
    grid1.LoadGrid cFileSave, flexFileData
    If grid1.rows > 1 Then cmdPrint.Enabled = True
    checkPhoto
End If
End Sub

Private Sub Command1_Click()
For I = 1 To grid1.rows - 1
    grid1.TextMatrix(I, grid1.Cols - 1) = 1
Next
End Sub

Private Sub cmdprintrep_Click()
Set PrintGrdNew.myForm = Me
Dim I As Long
For I = 1 To grid1.rows - 1
    If Not validPhoto(RetPhotoh(grid1.TextMatrix(I, 0))) Then
        grid1.RowHidden(I) = True
    End If
Next
grid1.ColHidden(grid1.Cols - 1) = True
PrintGrdNew.doprint grid1, 0.8, -3, "ÿ»«⁄… «·«⁄÷«¡", , , , False, False, 9
PrintGrdNew.Show 1
grid1.ColHidden(grid1.Cols - 1) = False
For I = 1 To grid1.rows - 1
    grid1.RowHidden(I) = False
Next
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1.LoadGrid cFileSave2, flexFileData
    For I = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(I, grid1.Cols - 1)) = 0 Then
            grid1.RemoveItem I
        End If
    Next
    cmdPrint.Enabled = (grid1.rows > 1)
    checkPhoto
End If
End Sub
Private Sub Form_Load()
openCon con
cFileSave = App.Path & "\" & Me.Name & ".grd"
cFileSave2 = App.Path & "\" & Me.Name & "_printed.grd"
Fixgrd
LoadText Me
xSeason.text = sSeason
xPrinted.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
Set cardqry_hfrm = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
grid1.SaveGrid cFileSave, flexFileData
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
grid1.SaveGrid cFileSave, flexFileData
End Sub
Private Sub Grid1_EnterCell()
grid1.Editable = IIf(grid1.Col = grid1.Cols - 1, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub Grid1_Keyup(KeyCode As Integer, Shift As Integer)
With grid1
If .rows = 1 Then Exit Sub
If KeyCode = 46 Then
    .RemoveItem grid1.Row
    .SaveGrid cFileSave, flexFileData
    CalcTotals
End If
End With
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub

Private Sub xCode1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub xCode1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim cWhere As String
    MemberH_LookupAll Me, oSearch
End If
End Sub

Private Sub xCode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub xCODE2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub
Private Function CountGrid() As Integer
With grid1
For I = 1 To grid1.rows - 1
    'If .TextMatrix(I, 6) = True Then CountGrid = CountGrid + 1
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
Private Sub SavePrinted()
With grid1
'Screen.MousePointer = 11

Dim aInsert As Variant, I As Long
con.BeginTrans
For I = 1 To .rows - 1
   If validPhoto(RetPhotoh(grid1.TextMatrix(I, 0))) And (Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0) Then
        aInsert = AddFlag(Empty, "MEMBER", grid1.TextMatrix(I, 0))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(I, 3)))
        
        aInsert = AddFlag(aInsert, "DATE", addDate(Now))
        aInsert = AddFlag(aInsert, "YEAR", addvalue(xSeason.text))
        
        con.Execute addInsert(aInsert, "file4_30")
        
        aInsert = AddFlag(Empty, "DATE_PRINT", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        con.Execute addUpdate(aInsert, "file3_10", "file3_10.CODE = " & grid1.TextMatrix(I, 0))
    End If
Next
con.CommitTrans
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
End Sub
Function eofGrd(cId) As Boolean
eofGrd = (grid1.FindRow(cId, , 0) = -1)
End Function
Private Function doprint() As Boolean
Dim nDiffer As Double
SettingArray(cUpMargin) = MyMeasure(0.4)
SettingArray(cRightMargin) = MyMeasure(-1.4)
SettingArray(cCardWidth) = MyMeasure(0)
SettingArray(cCardHeight) = MyMeasure(5.78)
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(8)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
nRows = SettingArray(cRows)
nspace = 0.62
nup = 0.3

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For I = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > nRows, 1, nRow)
    blastrow = (nRow = nRows)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡


prog1.Value = 0
prog1.Visible = True

For I = 1 To grid1.rows - 1
    If (validPhoto(RetPhotoh(grid1.TextMatrix(I, 0)))) And (Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0) Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 0.65
        
'        ' ‰Ê⁄ «·⁄÷ÊÌ…
'        tCard.AddNew
'        tCard!Right = MyMeasure(0.5)
'        tCard!Top = MyMeasure(0.1)
'        tCard!Width = MyMeasure(8)
'        tCard!Height = 0
'        tCard!FontName = "Arial"
'        tCard!FontBold = True
'        tCard!ForeColor = &HC0&
'        tCard!FontSize = 16
'        tCard!Text = TurnValue(IIf(grid1.TextMatrix(I, 6) = "1" Or grid1.TextMatrix(I, 6) = "", grid1.TextMatrix(I, 8), grid1.TextMatrix(I, 7)))
'        tCard!TextAlign = taCenterTop
'        tCard!CardNo = nCard
'        tCard.Update
        
        For i2 = 1 To 1
             ' «··ﬁ»
             
             tCard.AddNew
             tCard!Right = MyMeasure(0)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer)
             tCard!Width = MyMeasure(8)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 14
             If Check1.Value = 0 Then
                tCard!text = "œ⁄Ê…"
             End If
             tCard!TextAlign = taCenterTop
             tCard!CardNo = nCard
             tCard.Update
             
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 2)
             tCard!Width = MyMeasure(5.5)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 12
             tCard!text = TurnValue(IIf(Trim(grid1.TextMatrix(I, 3)) = "", ": «·«”„", grid1.TextMatrix(I, 3)))
             tCard!CardNo = nCard
             tCard.Update
             
             ' «”„ «·⁄÷Ê
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 3) - MyMeasure(0.1)
             tCard!Width = MyMeasure(8)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 12
             tCard!text = grid1.TextMatrix(I, 2)
             tCard!TextAlign = taRightTop
             tCard!CardNo = nCard
             tCard.Update
                                      
             
             ' —ﬁ„ «·⁄÷ÊÌ…
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 5)
             tCard!Width = MyMeasure(5)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 12
             tCard!text = ArbString("„”·”· —ﬁ„ :")
             tCard!TextAlign = taRightTop
             tCard!CardNo = nCard
             tCard.Update
            
             
             tCard.AddNew
             tCard!Right = MyMeasure(2.1) + MyMeasure(1.3)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 5)
             tCard!Width = MyMeasure(10)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 13
             tCard!text = ArbString(grid1.TextMatrix(I, 1))
             tCard!TextAlign = taRightTop
             tCard!CardNo = nCard
             tCard.Update
                           
             ' Ì‰ ÂÌ ›Ì
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 4)
             tCard!Width = 0
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 11
             tCard!text = "Ì‰ ÂÌ ›Ì " & Format(grid1.TextMatrix(I, 4), "yyyy/m/d")
             tCard!TextAlign = taRightTop
             tCard!CardNo = nCard
             tCard.Update
             
            tCard.AddNew
            tCard!Right = MyMeasure(6.7) - MyMeasure(0.8) + MyMeasure(0.2)
            tCard!Top = MyMeasure(4.2)
            tCard!Width = MyMeasure(3.1)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 10
            tCard!TextAlign = taCenterTop
            tCard!text = "ﬂ«» ‰/‰«’—«·‘«–·Ì"
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
            tCard!Top = MyMeasure(0.2) + MyMeasure(nDiffer * 6)
            tCard!Width = MyMeasure(9)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = False
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!text = "www.olympicclub.com"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            'ﬂ·„… —∆Ì” ‰«œÌ «·«”ﬂ‰œ—Ì…
            tCard.AddNew
            tCard!Right = MyMeasure(6) + MyMeasure(0.2)
            tCard!Top = MyMeasure(3.8)
            tCard!Width = MyMeasure(3)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 9
            tCard!TextAlign = taCenterTop
            tCard!text = "—∆Ì” „Ã·” «·«œ«—…"
            tCard!CardNo = nCard
            tCard.Update
        Next
        
' «·Ã“¡ «·«Ì”—
        
                
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.2) + MyMeasure(0.2)
        tCard!Top = MyMeasure(0.7)
        tCard!Width = MyMeasure(2.4) * 1.1
        tCard!Height = MyMeasure(2.8) * 1.1
        tCard!text = RetPhotoh(grid1.TextMatrix(I, 0))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
    End If
Next
prog1.Visible = False
tCard.Requery
doprint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Sub myProc()
ActiveControl.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 2)
Unload oSearch
End Sub
Private Sub checkPhoto()
Dim aPrint As Variant
With grid1
prog1.Value = 0
prog1.Visible = True
For I = 1 To grid1.rows - 1
    prog1.Value = Round(I / (grid1.rows - 1), 2) * 100
    aPrint = Printed_h(.TextMatrix(I, 0), xSeason.text, con)
    grid1.TextMatrix(I, 5) = myFormat_p(retFlag(aPrint, "date"))
    If IsDate(grid1.TextMatrix(I, 5)) Then
        .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &HE0E0E0
    End If
    If Not validPhoto(RetPhotoh(grid1.TextMatrix(I, 0))) Then grid1.Cell(flexcpForeColor, I, 0, I, .Cols - 1) = vbRed
Next
prog1.Visible = False
End With
End Sub
Private Sub myloadgrd()
Dim loctable As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, I As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT file3_10.*,file3_10.TITLE" & _
          " FROM file3_10 WHERE FILE3_10.CURRENT_MEM = 1"

If ValidNum(xCode1.text) Then
    cString = cString & turn(cString) & "file3_10.CODE  " & IIf(ValidNum(xcode2.text), " >= ", " = ") & xCode1.text
End If

If ValidNum(xcode2.text) Then
    cString = cString & turn(cString) & "file3_10.CODE <= " & xcode2.text
End If

If ValidNum(xNo1.text) Then
    cString = cString & turn(cString) & "file3_10.NO  " & IIf(ValidNum(xNo1.text), " >= ", " = ") & xNo1.text
End If

If ValidNum(xNo2.text) Then
    cString = cString & turn(cString) & "file3_10.NO <= " & xNo2.text
End If

If xPrinted.Value = 1 Then
    cString = cString & turn(cString) & "(dbo.f_printed_h(file3_10.code," & xSeason.text & ") IS NULL)"
End If

cString = cString & " ORDER BY file3_10.DESCA"

Set loctable = New ADODB.Recordset
With grid1
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    nRecordcount = loctable.RecordCount
End If
prog1.Visible = True
prog1.Value = 0
Dim nFound As Long
Do Until loctable.EOF
    I = I + 1
    prog1.Value = mRound(I / (nRecordcount)) * 100
    If .FindRow(loctable!CODE, , 0) = -1 Then
        .AddItem ""
        .TextMatrix(grid1.rows - 1, 0) = loctable!CODE
        .TextMatrix(grid1.rows - 1, 1) = loctable!NO & ""
        .TextMatrix(grid1.rows - 1, 2) = loctable!Desca & IIf(loctable!FAMILY, " Ê«·⁄«∆·…", "")
        .TextMatrix(grid1.rows - 1, 3) = loctable!Title & ""
        .TextMatrix(grid1.rows - 1, 4) = myFormat_p(loctable!DATE_END)
        .TextMatrix(grid1.rows - 1, 6) = loctable!FAMILY
        '.TextMatrix(grid1.rows - 1, 6) = myFormat_p(loctable!DATE_PRINT)
        .TextMatrix(grid1.rows - 1, .Cols - 1) = 1
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Me.MousePointer = 0
If grid1.rows > 1 Then
    grid1.Select 1, 0, 1, 1
    grid1.Sort = flexSortGenericAscending
End If
'CalcTotals
End With
End Sub
Private Sub Fixgrd()
With grid1
    .TextMatrix(0, 0) = "ﬂÊœ «·⁄÷Ê"
    .TextMatrix(0, 1) = "—ﬁ„ «·⁄÷Ê"
    .TextMatrix(0, 2) = "≈”„ «·⁄÷Ê"
    .TextMatrix(0, 3) = "«··ﬁ»"
    .TextMatrix(0, 4) = " «—ÌŒ ‰Â«Ì… «·ﬂ«—‰ÌÂ"
    .TextMatrix(0, 5) = " «—ÌŒ «Œ— ÿ»«⁄…"
    .TextMatrix(0, 6) = "Ê«·⁄«∆·…"
    .TextMatrix(0, 7) = "«Œ Ì«—"
    .ColDataType(6) = flexDTBoolean
    
    .ColDataType(.Cols - 1) = flexDTBoolean
    
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 4000
    .ColWidth(3) = 2000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    
    For I = 0 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    .ColDataType(0) = flexDTLong
    .ExplorerBar = flexExSortShow
End With
End Sub
Private Sub CalcTotals()
Dim nAll As Long, nPhoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(3).text = ""
StatusBar1.Panels(2).text = ""
StatusBar1.Panels(1).text = ""
If grid1.rows = 1 Then Exit Sub
For I = 0 To grid1.rows - 1
    nAll = nAll + 1
    If validPhoto(RetPhotoh(grid1.TextMatrix(I, 2))) Then nPhoto = nPhoto + 1
Next
nPhoto2 = ((Val(xRow.text) - 1) * 2) + (Val(xCol.text) - 1)
nPages = Fix(nPhoto / 10)
If nPhoto > 10 Then nLeft = nPhoto Mod 10
StatusBar1.Panels(3).text = "⁄œœ «·”Ã·«  : " & nAll
StatusBar1.Panels(2).text = "⁄œœ «·”Ã·«  »’Ê— : " & nPhoto
StatusBar1.Panels(1).text = "⁄œœ «·’›Õ«  : " & nPages
If nrest > 0 Then StatusBar1.Panels(3).text = StatusBar1.Panels(3).text & turn(StatusBar1.Panels(3).text, " ") & nrest & " ’Ê—…"
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub

Private Sub xNo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberH_LookupAll Me, oSearch
End If
End Sub

Private Sub xNo1_LostFocus()
myLostFocus xNo1
xcode_desca.Caption = ""
If Not ValidInt(xNo1.text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file3_10 where NO = " & xNo1.text)
If Not IsEmpty(aRet) Then
    xcode_desca.Caption = retFlag(aRet, "DESCA") & ""
End If
End Sub

Private Sub xSeason_GotFocus()
myGotFocus xSeason
End Sub
Private Sub xSeason_LostFocus()
myLostFocus xSeason
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xcode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xcode2
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xcode_desca.Caption = ""
If Not ValidInt(xCode1.text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file3_10 where code = " & xCode1.text)
If Not IsEmpty(aRet) Then
    xcode_desca.Caption = retFlag(aRet, "DESCA") & ""
End If
End Sub
Private Sub xDown_GotFocus()
myGotFocus xDown
End Sub
Private Sub xDown_LostFocus()
myLostFocus xDown
End Sub
Private Sub xRight_GotFocus()
myGotFocus xRight
End Sub
Private Sub xRight_LostFocus()
myLostFocus xRight
End Sub
Private Sub xCol_GotFocus()
myGotFocus xCol
End Sub
Private Sub xCol_LostFocus()
myLostFocus xCol
End Sub
Private Sub xRow_GotFocus()
myGotFocus xRow
End Sub
Private Sub xRow_LostFocus()
myLostFocus xRow
End Sub
Private Function myRev(pString As String) As String
Dim aString As Variant
If Trim(pString) = "" Then Exit Function
aString = Split(Trim(pString), "-")
If UBound(aString) = 0 Then
    myRev = Trim(aString(0))
Else
    myRev = Trim(aString(1)) & "-" & Trim(aString(0))
End If
End Function



