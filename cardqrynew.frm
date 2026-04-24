VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cardqryfrm 
   Caption         =   "ÿ»«⁄… «·ﬂ«—‰ÌÂ« "
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18600
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
   ScaleHeight     =   9750
   ScaleWidth      =   18600
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkMember 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·⁄÷Ê «·«”«”Ì ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9495
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   810
      Width           =   2580
   End
   Begin VB.CheckBox chkNoPhoto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   " √Œ›«¡ «⁄÷«¡ »œÊ‰ ’Ê—"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   900
      Width           =   2670
   End
   Begin VB.Frame Frame9 
      Height          =   780
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1125
      Width           =   1770
      Begin Threed.SSCommand cmdCsv 
         Height          =   555
         Left            =   45
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
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
         Picture         =   "cardqrynew.frx":0000
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "cardqrynew.frx":22B5
      End
   End
   Begin VB.CheckBox chkFawry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«⁄÷«¡ ›Ê—Ì ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   13815
      RightToLeft     =   -1  'True
      TabIndex        =   15
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
         TabIndex        =   19
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton cmdSavePrint 
         Caption         =   " „  «·ÿ»«⁄…"
         Height          =   390
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqrynew.frx":4438
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "cardqrynew.frx":68A4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
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
      TabIndex        =   31
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   35
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
         TabIndex        =   34
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
      TabIndex        =   26
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10575
      RightToLeft     =   -1  'True
      TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   13
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
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1170
      Width           =   6225
      Begin VB.CommandButton cmdAdd 
         Height          =   555
         Left            =   4815
         Picture         =   "cardqrynew.frx":913E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "cardqrynew.frx":B630
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   3465
         Picture         =   "cardqrynew.frx":DE1B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1320
      End
      Begin Threed.SSCommand cmdprintrep 
         Height          =   555
         Left            =   1620
         TabIndex        =   10
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
         Picture         =   "cardqrynew.frx":10245
         Caption         =   "ÿ»«⁄…  ﬁ—Ì—"
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "„Ê”„ «·ÿ»«⁄…"
      Enabled         =   0   'False
      Height          =   780
      Left            =   1935
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1125
      Visible         =   0   'False
      Width           =   1905
      Begin VB.TextBox xSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·–Ì ·„ Ìÿ»⁄ ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   810
      Width           =   2265
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   9375
      Width           =   18600
      _ExtentX        =   32808
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
      Left            =   -630
      Top             =   -225
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
      Left            =   -135
      Top             =   -225
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
      Left            =   -1620
      Top             =   45
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
      Left            =   -495
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
      Left            =   -1530
      Top             =   90
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
      Left            =   -1620
      Top             =   0
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
      Height          =   1365
      Left            =   12330
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   630
      Width           =   6135
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox xCode2 
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
      Begin VB.TextBox xAppend 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   945
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "«·ﬂÊœ"
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   225
         Width           =   3840
      End
      Begin VB.Label Label3 
         Caption         =   " ⁄÷Ê  «»⁄ "
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label10 
         Caption         =   "≈·Ì"
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
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         TabIndex        =   21
         Top             =   990
         Width           =   1005
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "cardqrynew.frx":12849
      Height          =   6090
      Left            =   90
      TabIndex        =   38
      Top             =   1980
      Width           =   18375
      _cx             =   32411
      _cy             =   10742
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
      Cols            =   14
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
      TabIndex        =   36
      Top             =   8100
      Width           =   2760
      Begin VB.CommandButton cmdSelect 
         BackColor       =   &H000000C0&
         Caption         =   "«Œ Ì«— «·ﬂ·"
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   2670
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   42
      Top             =   9135
      Visible         =   0   'False
      Width           =   18600
      _ExtentX        =   32808
      _ExtentY        =   423
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdAddUnpdate 
      Height          =   600
      Left            =   3870
      TabIndex        =   45
      Top             =   1305
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1058
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«÷«›… «⁄÷«¡ „”œœÌ‰"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "cardqryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileSave As String, cFileSave2 As String, cFilePrint As String
Dim oSearch As New Search
Dim con As New ADODB.Connection
Dim printTable As New ADODB.Recordset

Private Sub chkNoPhoto_Click()
HideNoPhoto
End Sub

Private Sub cmdAdd_Click()
'checkErr
myloadgrd
On Error Resume Next
grid1.SaveGrid cFileSave, flexFileData
Err.Clear
cmdPrint.Enabled = (grid1.rows > 1)
checkPhoto
If chkNoPhoto.Value = 1 Then HideNoPhoto
End Sub

Private Sub cmdCsv_Click()
If MsgBox("”Õ» ÿ»«⁄… «·ﬂ«—‰ÌÂ« ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If GetFile Then Inform " „ ‰ﬁ· «·»Ì«‰«  «·„ÿ»Ê⁄… »‰Ã«Õ"
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
For i = 1 To grid1.rows - 1
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.RowHidden(i) = True
    End If
Next
ToFileExel grid1, , , , , 0.9
For i = 1 To grid1.rows - 1
    grid1.RowHidden(i) = False
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

If Not doPrint Then
    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
    Exit Sub
End If

Set CardPrintNew.myForm = Me
CardPrintNew.PrintArray
CardPrintNew.Show 1

grid1.SaveGrid cFileSave2, flexFileData
If MsgBox(" „  «·ÿ»«⁄… ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    SavePrinted
    For i = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 Then
            grid1.RemoveItem i
        End If
    Next
    checkPhoto
End If
'grid1.SaveGrid cFilesave, flexFileData
End Sub
Private Sub CmdExit_Click()
Unload Me
Set cardqryfrm = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdClear_Click()
grid1.rows = 1
End Sub
Private Sub cmdMember_Click()
member.Show 1
End Sub
Private Sub CmdLastFillGrd_Click()
Dim fs As New FileSystemObject
If fs.FileExists(cFileSave) Then
    grid1.LoadGrid cFileSave, flexFileData
    If grid1.rows > 1 Then cmdPrint.Enabled = True
    checkPhoto
    CalcTotals
End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmdprintrep_Click()
Set PrintGrdNew.myForm = Me
Dim i As Long
For i = 1 To grid1.rows - 1
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 0))) Then
        grid1.RowHidden(i) = True
    End If
Next
grid1.ColHidden(grid1.Cols - 1) = True
PrintGrdNew.doPrint grid1, 0.8, -3, "ÿ»«⁄… «·«⁄÷«¡", , , , False, False, 9
PrintGrdNew.Show 1
grid1.ColHidden(grid1.Cols - 1) = False
For i = 1 To grid1.rows - 1
    grid1.RowHidden(i) = False
Next
End Sub

Private Sub cmdSelect_Click()
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, grid1.Cols - 1) = 1
Next
grid1.Cell(flexcpChecked, 0, 12) = 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1.LoadGrid cFileSave2, flexFileData
    For i = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(i, grid1.Cols - 1)) = 0 Then
            grid1.RemoveItem i
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
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
    If Row = 0 And Col = .Cols - 1 Then
        For i = 1 To .rows - 1
            .TextMatrix(i, .Cols - 1) = IIf(.Cell(flexcpChecked, 0, Col) = 1, -1, 0)
        Next
    End If
    .SaveGrid cFileSave, flexFileData
End With
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
grid1.SaveGrid cFileSave, flexFileData
End Sub
Private Sub Grid1_EnterCell()
grid1.Editable = IIf(grid1.Col = grid1.Cols - 1 Or grid1.Col = 12, flexEDKbdMouse, flexEDNone)
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
Private Sub xCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    xUnCode.Caption = ""
    If xCode.text = "" Then Exit Sub
    xUnCode.ForeColor = -2147483630
    If Val(unMyCodeBar(xCode.text, "1")) <> 1 Then
        xUnCode.Caption = "Error"
        xUnCode.ForeColor = vbRed
    Else
        xUnCode.Caption = unMyCodeBar(xCode.text)
    End If
    myGotFocus xCode
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub

Private Sub cmdAddUnpdate_Click()
myLoadGrdPaid
On Error Resume Next
grid1.SaveGrid cFileSave, flexFileData
Err.Clear
cmdPrint.Enabled = (grid1.rows > 1)
checkPhoto True
grid1.Cell(flexcpChecked, 0, grid1.Cols - 1) = 1
End Sub

Private Sub xCode1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub xCode1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim cWhere As String
    MemberLookupAll Me, oSearch
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
For i = 1 To grid1.rows - 1
    'If .TextMatrix(I, 6) = True Then CountGrid = CountGrid + 1
    CountGrid = CountGrid + 1
Next
End With
End Function
Private Sub countPrint()
nCountPrint = 0
With grid1
For i = 1 To .rows - 1
   If .TextMatrix(i, 6) = True Then nCountPrint = nCountPrint + 1
Next
lblCount.Caption = nCountPrint / 1
End With
End Sub
Private Function MakeString()
MakeString = "#" & ";"
MakeString = MakeString & "|#" & 0 & ";" & "ﬂ«—‰ÌÂ ÃœÌœ"
MakeString = MakeString & "|#" & 1 & ";" & "»œ· ›«ﬁœ"
End Function
Private Sub SavePrinted()
With grid1
'Screen.MousePointer = 11
dTime = Time
dDate = Date

Dim aInsert As Variant, i As Long
con.BeginTrans
For i = 1 To .rows - 1
   If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And (Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0) Then
        aInsert = AddFlag(Empty, "MEMBER", grid1.TextMatrix(i, 0))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DESCA2", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(i, 5)))
        
        aInsert = AddFlag(aInsert, "RELATION", addstring(grid1.TextMatrix(i, 5 + 1)))
        aInsert = AddFlag(aInsert, "RELATION_DESCA", addstring(grid1.TextMatrix(i, 6 + 1)))
        aInsert = AddFlag(aInsert, "TYPE_DESCA", addstring(grid1.TextMatrix(i, 7 + 1)))
        
        aInsert = AddFlag(aInsert, "DATE", addDate(Now))
        aInsert = AddFlag(aInsert, "YEAR", addvalue(xSeason.text))
        
        con.Execute addInsert(aInsert, "file4_10")
        
        aInsert = AddFlag(Empty, "DATE_PRINT", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        If grid1.TextMatrix(i, 1) = "" Then
            con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & grid1.TextMatrix(i, 0))
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "FILE1_11.MEMBER = " & grid1.TextMatrix(i, 0) & " AND FILE1_11.CODE = " & grid1.TextMatrix(i, 1))
        End If
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
Private Function doPrint2024() As Boolean
Dim nDiffer As Double
SettingArray(cUpMargin) = MyMeasure(0.4)
SettingArray(cRightMargin) = MyMeasure(-1.45)
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
For i = 1 To nBegin - 1
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

For i = 1 To grid1.rows - 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And (Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0) Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 0.45
                
        For i2 = 1 To 1
             
                ' «”„ «·⁄÷Ê
            tCard.AddNew
            tCard!Right = MyMeasure(0)
            tCard!Top = MyMeasure(1.15)
            tCard!Width = MyMeasure(8)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!text = TurnValue(IIf(grid1.TextMatrix(i, 6) = "1" Or grid1.TextMatrix(i, 6) = "", grid1.TextMatrix(i, 8), grid1.TextMatrix(i, 7)))
            If grid1.ValueMatrix(i, 12) <> 0 Then
                tCard!FontSize = 14
                tCard!text = ArbString(tCard!text & " (»œ· ›«ﬁœ)")
            Else
                tCard!FontSize = 16
            End If
            tCard!TextAlign = taCenterTop
            tCard!CardNo = nCard
            tCard.Update
             
             ' «··ﬁ»
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.5) + MyMeasure(nDiffer * 3)
             tCard!Width = MyMeasure(6)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 13
             tCard!text = TurnValue(ArbString(IIf(Trim(grid1.TextMatrix(i, 5)) = "", "«·«”„ :", grid1.TextMatrix(i, 5) & " : ")))
             'tCard!text = ArbString(tCard!text & " " & IIf(grid1.TextMatrix(I, 1) = "", grid1.TextMatrix(I, 3), grid1.TextMatrix(I, 4)))
             tCard!CardNo = nCard
             tCard.Update
                                      
                                      
             ' «··ﬁ»
             tCard.AddNew
             tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
             tCard!Top = MyMeasure(0.55) + MyMeasure(nDiffer * 4) - MyMeasure(0.02)
             tCard!Width = MyMeasure(5.5)
             tCard!Height = 0
             tCard!FontName = "Arial"
             tCard!FontBold = True
             tCard!ForeColor = vbBlack
             tCard!FontSize = 12
             tCard!text = IIf(grid1.TextMatrix(i, 1) = "", grid1.TextMatrix(i, 3), grid1.TextMatrix(i, 4))
             tCard!CardNo = nCard
             tCard.Update
                                      
             
            ' —ﬁ„ «·⁄÷ÊÌ…
            tCard.AddNew
            tCard!Right = MyMeasure(0.2) + MyMeasure(1.2)
            tCard!Top = MyMeasure(0.65) + MyMeasure(nDiffer * 5)
            tCard!Width = MyMeasure(5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 13
            tCard!text = ArbString("—ﬁ„ «·⁄÷ÊÌ… :")
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
                
            tCard.AddNew
            tCard!Right = MyMeasure(2.3) + MyMeasure(1.5)
            tCard!Top = MyMeasure(0.65) + MyMeasure(nDiffer * 5)
            tCard!Width = MyMeasure(10)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 13
            'tCard!text = ArbString(grid1.TextMatrix(i, 1) & IIf(grid1.TextMatrix(i, 1) <> "", "-", "") & grid1.TextMatrix(i, 0))
            tCard!text = grid1.TextMatrix(i, 0) & IIf(grid1.TextMatrix(i, 1) <> "", "-", "") & grid1.TextMatrix(i, 1)
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
                
'                 Ì‰ ÂÌ ›Ì
            tCard.AddNew
            tCard!Right = MyMeasure(2.3)
            tCard!Top = MyMeasure(0.9) + MyMeasure(nDiffer * 6)
            tCard!Width = MyMeasure(8)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!text = "Ì‰ ÂÌ ›Ì "
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
                                                                              
'                 Ì‰ ÂÌ ›Ì
            tCard.AddNew
            tCard!Right = MyMeasure(3.8)
            tCard!Top = MyMeasure(0.9) + MyMeasure(nDiffer * 6)
            tCard!Width = MyMeasure(8)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!text = myFormat_p("30/06/2025")
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
                                                                              
            tCard.AddNew
            tCard!Top = MyMeasure(1.1) + MyMeasure(nDiffer * 7)
            tCard!Right = MyMeasure(0.2)
            tCard!Width = MyMeasure(8)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 10
            tCard!text = "www.olympicclub.com"
            tCard!TextAlign = taCenterTop
            tCard!CardNo = nCard
            tCard.Update
                        
            'ﬂ·„… —∆Ì” ‰«œÌ «·«”ﬂ‰œ—Ì…
            tCard.AddNew
            tCard!Right = MyMeasure(6) + MyMeasure(0.2)
            tCard!Top = MyMeasure(3.85)
            tCard!Width = MyMeasure(3)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!TextAlign = taCenterTop
            tCard!text = "—∆Ì” „Ã·” «·«œ«—…"
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(6.7) - MyMeasure(0.8) + MyMeasure(0.3)
            tCard!Top = MyMeasure(4.3)
            tCard!Width = MyMeasure(3.1)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!TextAlign = taCenterTop
            tCard!text = "‰«’— ‰»Ì· «·‘«–·Ï"
            tCard!CardNo = nCard
            tCard.Update
            
            
        Next
        
' «·Ã“¡ «·«Ì”—
        
                
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.4) + MyMeasure(0)
        tCard!Top = MyMeasure(0.7)
        tCard!Width = MyMeasure(2.4) * 1.1
        tCard!Height = MyMeasure(2.8) * 1.1
        tCard!text = retPhoto(grid1.TextMatrix(i, 2))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
    End If
Next
prog1.Visible = False
tCard.Requery
doPrint2024 = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Private Function doPrint() As Boolean
Dim nDiffer As Double
SettingArray(cUpMargin) = MyMeasure(0.4)
SettingArray(cRightMargin) = MyMeasure(-1.25)
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
For i = 1 To nBegin - 1
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

For i = 1 To grid1.rows - 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And (Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0) Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 0.6
                                        
        If grid1.ValueMatrix(i, 12) <> 0 Then
            ' «··ﬁ»
            tCard.AddNew
            tCard!Right = MyMeasure(1.7)
            tCard!Top = MyMeasure(0.7)
            tCard!Width = MyMeasure(3)
            tCard!Height = 0
            tCard!FontName = "GE SS two Medium"
            tCard!FontBold = False
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!text = "»œ· ›‹«ﬁ‹œ"
            tCard!TextAlign = taRightTop
            'tCard!text = ArbString(tCard!text & " " & IIf(grid1.TextMatrix(I, 1) = "", grid1.TextMatrix(I, 3), grid1.TextMatrix(I, 4)))
            tCard!CardNo = nCard
            tCard.Update
        End If
                                        
                                        
        ' «··ﬁ»
        tCard.AddNew
        tCard!Right = MyMeasure(1.7)
        tCard!Top = MyMeasure(1.7)
        tCard!Width = MyMeasure(3)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!FontSize = 9
        'tCard!text = grid1.TextMatrix(i, 5)
        tCard!TextAlign = taRightTop
        'tCard!text = ArbString(tCard!text & " " & IIf(grid1.TextMatrix(I, 1) = "", grid1.TextMatrix(I, 3), grid1.TextMatrix(I, 4)))
        tCard!CardNo = nCard
        tCard.Update
                                        
        ' «··ﬁ»
        tCard.AddNew
        tCard!Right = MyMeasure(1.7)
        tCard!Top = MyMeasure(1.63)
        tCard!Width = MyMeasure(3)
        tCard!Height = 0
        tCard!FontName = "Mohammad Head"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 11
        tCard!text = grid1.TextMatrix(i, 5)
        tCard!TextAlign = taRightTop
        'tCard!text = ArbString(tCard!text & " " & IIf(grid1.TextMatrix(I, 1) = "", grid1.TextMatrix(I, 3), grid1.TextMatrix(I, 4)))
        tCard!CardNo = nCard
        tCard.Update
                                        
        ' «·«”„
        tCard.AddNew
        tCard!Right = MyMeasure(2.7)
        tCard!Top = MyMeasure(1.6) + MyMeasure(nDiffer * 1)
        tCard!Width = MyMeasure(5.5)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!FontSize = 9
       ' tCard!text = IIf(grid1.TextMatrix(i, 1) = "", grid1.TextMatrix(i, 3), grid1.TextMatrix(i, 4))
        tCard!CardNo = nCard
        tCard.Update
                                 
        ' «·«”„
        tCard.AddNew
        tCard!Right = MyMeasure(2.7)
        tCard!Top = MyMeasure(1.55) + MyMeasure(nDiffer * 1)
        tCard!Width = MyMeasure(5.5)
        tCard!Height = 0
        tCard!FontName = "Mohammad Head"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 11
        tCard!text = IIf(grid1.TextMatrix(i, 1) = "", grid1.TextMatrix(i, 3), grid1.TextMatrix(i, 4))
        tCard!CardNo = nCard
        tCard.Update
                                 
        tCard.AddNew
        tCard!Right = MyMeasure(2.7)
        tCard!Top = MyMeasure(1.6) + MyMeasure(nDiffer * 2)
        tCard!Width = MyMeasure(10)
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!FontSize = 11
        tCard!text = grid1.TextMatrix(i, 0) & IIf(grid1.TextMatrix(i, 1) <> "", "-", "") & grid1.TextMatrix(i, 1)
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
                                      
        ' ‰Ê⁄ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(2.7)
        tCard!Top = MyMeasure(1.6) + MyMeasure(nDiffer * 3)
        tCard!Width = MyMeasure(8)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
       ' tCard!text = TurnValue(IIf(grid1.TextMatrix(i, 6) = "1" Or grid1.TextMatrix(i, 6) = "", grid1.TextMatrix(i, 8), grid1.TextMatrix(i, 7)))
'        If grid1.ValueMatrix(i, 12) <> 0 Then
'            tCard!FontSize = 8
'            tCard!text = ArbString(tCard!text & " (»œ· ›«ﬁœ)")
'        Else
            tCard!FontSize = 9
'        End If
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
                                                                                    
        ' ‰Ê⁄ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(2.7)
        tCard!Top = MyMeasure(1.45) + MyMeasure(nDiffer * 3)
        tCard!Width = MyMeasure(8)
        tCard!Height = 0
        tCard!FontName = "Mohammad Head"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!text = TurnValue(IIf(grid1.TextMatrix(i, 6) = "1" Or grid1.TextMatrix(i, 6) = "", grid1.TextMatrix(i, 8), grid1.TextMatrix(i, 7)))
'        If grid1.ValueMatrix(i, 12) <> 0 Then
'            tCard!FontSize = 8
'            tCard!text = ArbString(tCard!text & " (»œ· ›«ﬁœ)")
'        Else
            tCard!FontSize = 11
'        End If
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
                                                                                    
' «·Ã“¡ «·«Ì”—
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.5) + MyMeasure(0)
        tCard!Top = MyMeasure(0.7)
        tCard!Width = MyMeasure(2.4) * 1.1
        tCard!Height = MyMeasure(2.8) * 1.1
        tCard!text = retPhoto(grid1.TextMatrix(i, 2))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
    End If
Next
prog1.Visible = False
tCard.Requery
doPrint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Sub myProc()
ActiveControl.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Private Sub checkPhoto(Optional bNoCheck As Boolean = False)
Dim aPrint As Variant
With grid1
prog1.Value = 0
prog1.Visible = True
For i = 1 To grid1.rows - 1
    prog1.Value = Round(i / (grid1.rows - 1), 2) * 100
    If bNoCheck Then
        aPrint = Printed(.TextMatrix(i, 0), .TextMatrix(i, 1), xSeason.text, con)
        grid1.TextMatrix(i, 11) = myFormat_p(retFlag(aPrint, "date"))
        If IsDate(grid1.TextMatrix(i, 11)) Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE0E0E0
        End If
    End If
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then grid1.Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
Next
prog1.Visible = False
End With
End Sub
Private Sub myloadgrd()
Dim loctable As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, i As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT FILE1_10.*,TYPE_CODES.DESCA AS TYPE_DESCA,FILE6_20H.FORM_NO as last_doc,FILE6_20H.DATE as last_date2,FILE1_10.TITLE" & _
          " FROM FILE1_10 LEFT JOIN TYPE_CODES ON FILE1_10.TYPE = TYPE_CODES.CODE " & _
          " INNER JOIN vw_last_paid_no_save on FILE1_10.CODE = [dbo].[vw_last_paid_no_save].CODE" & _
          " INNER JOIN FILE6_20H ON vw_last_paid_no_save.DOC_NO = FILE6_20H.DOC_NO"
cString = cString & " WHERE FILE1_10.DIED = 0"

If ValidInt(xCode1.text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.text), " >= ", " = ") & xCode1.text
End If

If ValidInt(xCode2.text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.text
End If

If IsDate(xDate1.text) Then
     cString = cString & turn(cString) & "FILE6_20H.DATE >= " & DateSq(xDate1.text)
End If

If IsDate(xDate2.text) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate2.text)
End If

If xPrinted.Value = 1 Then
    cString = cString & turn(cString) & "(dbo.f_printed(file1_10.code,NULL," & xSeason.text & ") IS NULL)"
End If

If chkFawry.Value = 1 Then
    cString = cString & turn(cString) & " FILE6_20H.ISFAWRY = 1"
End If

If xSeason.text <> "" Then
    cString = cString & turn(cString) & " [vw_last_paid_no_save].YEAR_CODE = " & xSeason.text
End If

cString = cString & " ORDER BY FILE1_10.DESCA"


Set loctable = New ADODB.Recordset
With grid1

Set loctable = myCmd(cString, con)
If Not (loctable.EOF And loctable.BOF) Then
    nRecordcount = loctable.RecordCount
End If
prog1.Visible = True
prog1.Value = 0

Dim nFound As Long
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    If .FindRow(loctable!code, , 0) = -1 Then
        .AddItem ""
        .TextMatrix(grid1.rows - 1, 0) = loctable!code
        .TextMatrix(grid1.rows - 1, 1) = ""
                                            
                                                                                        
        ' «·„⁄—Ê÷
        .TextMatrix(grid1.rows - 1, 2) = loctable!code
        .TextMatrix(grid1.rows - 1, 3) = loctable!Desca & ""
        .TextMatrix(grid1.rows - 1, 4) = ""
        .TextMatrix(grid1.rows - 1, 5) = loctable!Title & ""
        .TextMatrix(grid1.rows - 1, 6) = ""
        
        .TextMatrix(grid1.rows - 1, 7) = ""
        
        .TextMatrix(grid1.rows - 1, 8) = "⁄÷Ê ⁄«„·"
        .TextMatrix(grid1.rows - 1, 9) = loctable!last_doc & ""
        .TextMatrix(grid1.rows - 1, 10) = myFormat_p(loctable!last_date2)
        .TextMatrix(grid1.rows - 1, .Cols - 1) = -1
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing

If chkMember.Value = 0 Then
    cString = "SELECT file1_11.Code,File1_11.Relation,TYPE_CODES.DESCA AS TYPE_DESCA,File1_11.Member,File1_11.Title,file1_11.descA,FILE1_10.DESCA as Desca_Member,Relation_codes.Desca as Desca_relation,FILE6_20H.FORM_NO as last_doc,FILE6_20H.DATE as last_date2 " & _
              " From File1_11 Inner Join File1_10 On File1_11.Member = File1_10.Code " & _
              " INNER join relation_codes on file1_11.relation = relation_codes.code LEFT JOIN TYPE_CODES ON FILE1_10.TYPE = TYPE_CODES.CODE" & _
              " INNER JOIN vw_last_paid_no_save on FILE1_10.CODE = [dbo].[vw_last_paid_no_save].CODE" & _
              " INNER JOIN FILE6_20H ON vw_last_paid_no_save.DOC_NO = FILE6_20H.DOC_NO"
    'cString = cString & " WHERE FILE1_10.DIED = 0"
    'cString = cString & " WHERE FILE6_20H.TYPE <> 2"
              
    If xAppend.text = "" Then
        If ValidInt(xCode1.text) Then
            cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.text), " >= ", " = ") & xCode1.text
        End If
        
        If ValidInt(xCode2.text) Then
            cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.text
        End If
    Else
        cString = cString & turn(cString) & "File1_11.MEMBER = " & xCode1.text & " and File1_11.code = " & xAppend.text
    End If
    
    If IsDate(xDate1.text) Then
         cString = cString & turn(cString) & "FILE6_20H.DATE >= " & DateSq(xDate1.text)
    End If
    
    If IsDate(xDate2.text) Then
        cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate2.text)
    End If
    
    If chkFawry.Value = 1 Then
         cString = cString & turn(cString) & " FILE6_20H.ISFAWRY = 1"
    End If
    
    
    If xPrinted.Value = 1 Then
        cString = cString & turn(cString) & "(dbo.f_printed(file1_11.member,file1_11.code," & xSeason.text & ") IS NULL)"
    End If
    
    If xSeason.text <> "" Then
        cString = cString & turn(cString) & " [vw_last_paid_no_save].YEAR_CODE = " & xSeason.text
    End If
    
    Set loctable = New ADODB.Recordset
    Set loctable = myCmd(cString, con)
    If Not (loctable.EOF And loctable.BOF) Then
        nRecordcount = loctable.RecordCount
    End If
    prog1.Visible = True
    prog1.Value = 0
    
    i = 0
    Do Until loctable.EOF
        i = i + 1
        prog1.Value = Round(i / (nRecordcount), 2) * 100
        If .FindRow(loctable!member & "-" & loctable!code, , 2) = -1 Then
            .AddItem ""
            .TextMatrix(.rows - 1, 0) = loctable!member
            .TextMatrix(.rows - 1, 1) = loctable!code
             
             ' «·„⁄—Ê÷
            .TextMatrix(.rows - 1, 2) = loctable!member & "-" & loctable!code
            .TextMatrix(.rows - 1, 3) = loctable!Desca_Member & ""
            .TextMatrix(.rows - 1, 4) = loctable!Desca & ""
            .TextMatrix(.rows - 1, 5) = loctable!Title & ""
            
            .TextMatrix(.rows - 1, 6) = loctable!RELATION & ""
            .TextMatrix(.rows - 1, 7) = "⁄÷Ê  «»⁄"
            .TextMatrix(grid1.rows - 1, 8) = "⁄÷Ê ⁄«„·"
            .TextMatrix(grid1.rows - 1, 9) = loctable!last_doc & ""
            .TextMatrix(grid1.rows - 1, 10) = myFormat_p(loctable!last_date2)
            .TextMatrix(grid1.rows - 1, .Cols - 1) = -1
        End If
        loctable.MoveNext
    Loop
    loctable.Close
    Set loctable = Nothing
End If
prog1.Visible = False
Me.MousePointer = 0
If grid1.rows > 1 Then
    grid1.Select 1, 0, 1, 1
    grid1.Sort = flexSortGenericAscending
End If
CalcTotals
End With
End Sub
Private Sub myLoadGrdPaid()
Dim loctable As ADODB.Recordset, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, i As Long, bAddRow As Boolean
Me.MousePointer = 11


Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", xSeason.text)

If IsNumeric(xCode1.text) And Not IsNumeric(xCode2.text) Then
    aPrm = AddFlag(aPrm, "CODE", xCode1.text)
ElseIf IsNumeric(xCode1.text) Then
    aPrm = AddFlag(aPrm, "CODE1", xCode1.text)
End If

If IsNumeric(xCode2.text) Then
    aPrm = AddFlag(aPrm, "CODE2", xCode2.text)
End If

Set loctable = New ADODB.Recordset
With grid1

Set loctable = myCmd("[dbo].[sp_paid_no_card]", con, adStoredProc, aPrm, 1000)
If Not (loctable.EOF And loctable.BOF) Then
    nRecordcount = loctable.RecordCount
End If
prog1.Visible = True
prog1.Value = 0
grid1.rows = 1

Dim nFound As Long
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    
    If IsNull(loctable!code) Then
        .AddItem ""
        .TextMatrix(grid1.rows - 1, 0) = loctable!member
        .TextMatrix(grid1.rows - 1, 1) = ""
                                            
                                                                                        
        ' «·„⁄—Ê÷
        .TextMatrix(grid1.rows - 1, 2) = loctable!member
        .TextMatrix(grid1.rows - 1, 3) = loctable!Desca & ""
        .TextMatrix(grid1.rows - 1, 4) = ""
        .TextMatrix(grid1.rows - 1, 5) = loctable!Title & ""
        .TextMatrix(grid1.rows - 1, 6) = ""
        
        .TextMatrix(grid1.rows - 1, 7) = ""
        
        .TextMatrix(grid1.rows - 1, 8) = "⁄÷Ê ⁄«„·"
        .TextMatrix(grid1.rows - 1, 9) = loctable!doc_no & ""
        .TextMatrix(grid1.rows - 1, 10) = myFormat_p(loctable!Date)
        .TextMatrix(grid1.rows - 1, .Cols - 1) = -1
    Else
        .AddItem ""
        .TextMatrix(.rows - 1, 0) = loctable!member
        .TextMatrix(.rows - 1, 1) = loctable!code
         
         ' «·„⁄—Ê÷
        .TextMatrix(.rows - 1, 2) = loctable!member & "-" & loctable!code
        .TextMatrix(.rows - 1, 3) = loctable!Desca_Member
        .TextMatrix(.rows - 1, 4) = loctable!Desca & ""
        .TextMatrix(.rows - 1, 5) = loctable!Title & ""
        
        .TextMatrix(.rows - 1, 6) = loctable!RELATION & ""
        .TextMatrix(.rows - 1, 7) = "⁄÷Ê  «»⁄"
        .TextMatrix(grid1.rows - 1, 8) = "⁄÷Ê ⁄«„·"
        .TextMatrix(grid1.rows - 1, 9) = loctable!doc_no & ""
        .TextMatrix(grid1.rows - 1, 10) = myFormat_p(loctable!Date)
        .TextMatrix(grid1.rows - 1, .Cols - 1) = -1
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
End With
End Sub
Private Sub myloadgrdChair()
Dim loctable As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, i As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT FILE1_10.* FROM FILE1_10 WHERE file1_10.CHAIR = 1"

If ValidInt(xCode1.text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.text), " >= ", " = ") & xCode1.text
End If

If ValidInt(xCode2.text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.text
End If

cString = cString & " ORDER BY FILE1_10.DESCA"

Set loctable = New ADODB.Recordset
With grid1
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    nRecordcount = loctable.RecordCount
End If
prog1.Visible = True
prog1.Value = 0

If chkChair.Value = 1 Then
    grid1.rows = 1
End If

Dim nFound As Long
Do Until loctable.EOF
    i = i + 1
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    If .FindRow(loctable!code, , 0) = -1 Then
        .AddItem ""
        .TextMatrix(grid1.rows - 1, 0) = loctable!code
        .TextMatrix(grid1.rows - 1, 1) = ""
                                            
        ' «·„⁄—Ê÷
        .TextMatrix(grid1.rows - 1, 2) = loctable!code
        .TextMatrix(grid1.rows - 1, 3) = loctable!Desca & ""
        .TextMatrix(grid1.rows - 1, 4) = ""
        .TextMatrix(grid1.rows - 1, 5) = loctable!Title & ""
        .TextMatrix(grid1.rows - 1, 6) = ""
        
        .TextMatrix(grid1.rows - 1, 7) = ""
        
        '.TextMatrix(grid1.rows - 1, 8) = loctable!type_desca & ""
        .TextMatrix(grid1.rows - 1, 8) = loctable!notes
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
End With
End Sub
Private Sub Fixgrd()
With grid1
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "≈”„ «·⁄÷Ê"
    
    .TextMatrix(0, 4) = "«·⁄÷Ê «· «»⁄"
    .TextMatrix(0, 5) = "«··ﬁ»"
    
    .TextMatrix(0, 6) = "œ—Ã… «·ﬁ—«»…"
    .TextMatrix(0, 7) = "œ—Ã… «·ﬁ—«»…"
    
    .TextMatrix(0, 8) = "‰Ê⁄ «·⁄÷ÊÌ…"
    .TextMatrix(0, 9) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 10) = " «—ÌŒ «·„” ‰œ"
    
    .TextMatrix(0, 11) = " «—ÌŒ «Œ— ÿ»«⁄…"
    .TextMatrix(0, 12) = "»œ· ›«ﬁœ"
    .TextMatrix(0, 13) = "«Œ Ì«—"
    
    .ColDataType(12) = flexDTBoolean
    .ColDataType(.Cols - 1) = flexDTBoolean
    
    .ColHidden(0) = True
    .ColHidden(1) = True
    .ColHidden(6) = True
    
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 2500
    
    .ColWidth(5 + 1) = 1500
    .ColWidth(6 + 1) = 1500
    .ColWidth(7 + 1) = 2000
    .ColWidth(8 + 1) = 1500
    .ColWidth(9 + 1) = 1500
    .ColWidth(10 + 1) = 1400
    .ColWidth(11 + 1) = 1000
    
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColDataType(0) = flexDTLong
    .ColDataType(1) = flexDTLong
    '.ExplorerBar = flexExSortShow
    .Cell(flexcpChecked, 0, .Cols - 1, 0, .Cols - 1) = 2

End With
End Sub
Private Sub CalcTotals()
Dim nAll As Long, nPhoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(3).text = ""
StatusBar1.Panels(2).text = ""
StatusBar1.Panels(1).text = ""
If grid1.rows = 1 Then Exit Sub
For i = 1 To grid1.rows - 1
    nAll = nAll + 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then nPhoto = nPhoto + 1
Next
'nPhoto2 = ((Val(xRow.text) - 1) * 2) + (Val(xCol.text) - 1)
'nPages = Fix(nPhoto / 10)
If nPhoto > 10 Then nLeft = nPhoto Mod 10
StatusBar1.Panels(3).text = "⁄œœ «·ﬂ«—‰ÌÂ«  : " & nAll
StatusBar1.Panels(2).text = "⁄œœ «·ﬂ«—‰ÌÂ«  »’Ê— : " & nPhoto
StatusBar1.Panels(1).text = "⁄œœ «·ﬂ«—‰ÌÂ«  »œÊ‰ ’Ê— : " & nAll - nPhoto
'If nrest > 0 Then StatusBar1.Panels(3).text = StatusBar1.Panels(3).text & turn(StatusBar1.Panels(3).text, " ") & nrest & " ’Ê—…"
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

Private Sub xSeason_GotFocus()
myGotFocus xSeason
End Sub
Private Sub xSeason_LostFocus()
myLostFocus xSeason
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
Private Sub xAppend_GotFocus()
myGotFocus xAppend
End Sub
Private Sub xAppend_LostFocus()
myLostFocus xAppend
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xcode_desca.Caption = ""
If Not ValidInt(xCode1.text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode1.text)
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
Private Sub DelUncheck()
Dim i As Long
For i = grid1.rows - 1 To 1 Step -1
    If mRound(grid1.TextMatrix(i, grid1.Cols - 1)) = 0 Then
        grid1.RemoveItem i
    End If
Next
End Sub
Private Function GetFile() As Boolean
Dim cSv As New ChilkatCsv

Dim cFile As String
Dim cString As String

cFile = App.Path & "\csv\card_printed2024.csv"

If Dir(cFile) = "" Then
    MsgBox "«·„·› €Ì— „ÊÃÊœ"
    Exit Function
End If

nAccess = cSv.LoadFile(cFile)

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If

If Dir(cFile) = "" Then
    MsgBox "«·„·› €Ì— „ÊÃÊœ"
    Exit Function
End If

nAccess = cSv.LoadFile(cFile)

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If

Dim aInsert As Variant
'con.BeginTrans
'On Error GoTo myerror

Dim sMember As String
Dim sCode As String
Dim sDesca As String
Dim sdesca2 As String
Dim stitle As String
Dim sType_desca As String
Dim aRelation As Variant
Dim sRelation
Dim sRelation_desca As String

'con.BeginTrans
Dim cmd As ADODB.Command
Set cmd = myCommand("delete from file4_10 where [year] = " & sSeason, con)

'con.Execute "delete from file4_10 where [year] = " & sSeason
prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption
For i = 1 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & cSv.NumRows - 1
        
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    sMember = mySplit(cSv.GetCell(i, 0), 1, "-")
    sCode = mySplit(cSv.GetCell(i, 0), 2, "-")
    'sDesca = Member_Load(sMember, "desca", con) & ""
    If sCode <> "" Then
        sdesca2 = cSv.GetCell(i, 2)
    Else
        sdesca2 = ""
    End If
    stitle = cSv.GetCell(i, 1)
    sType_desca = cSv.GetCell(i, 3)
    
    If sMember <> "" Then
        aInsert = AddFlag(Empty, "MEMBER", sMember)
        aInsert = AddFlag(aInsert, "CODE", addvalue(sCode))
        aInsert = AddFlag(aInsert, "DESCA", addstring(sDesca))
        aInsert = AddFlag(aInsert, "DESCA2", addstring(sdesca2))
        aInsert = AddFlag(aInsert, "TITLE", addstring(stitle))
        
        aInsert = AddFlag(aInsert, "RELATION", addstring(sRelation))
        aInsert = AddFlag(aInsert, "RELATION_DESCA", addstring(sRelation_desca))
        aInsert = AddFlag(aInsert, "TYPE_DESCA", addstring(sType_desca))
        
        aInsert = AddFlag(aInsert, "DATE", addDate(Now))
        aInsert = AddFlag(aInsert, "YEAR", sSeason)
        
        con.Execute addInsert(aInsert, "file4_10")
        
        aInsert = AddFlag(Empty, "DATE_PRINT", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        If sCode = "" Then
            con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & sMember)
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "FILE1_11.MEMBER = " & sMember & " AND FILE1_11.CODE = " & sCode)
        End If
    End If
Next

con.Execute "UPDATE FILE4_10 " & _
            " SET FILE4_10.RELATION = FILE1_11.RELATION," & _
            " FILE4_10.RELATION_DESCA = RELATION_CODES.DESCA" & _
            " FROM FILE4_10 INNER JOIN FILE1_11 ON FILE4_10.MEMBER = FILE1_11.MEMBER AND FILE4_10.CODE = FILE1_11.CODE " & _
            " LEFT JOIN  RELATION_CODES ON FILE1_11.RELATION = RELATION_CODES.CODE" & _
            " WHERE FILE4_10.YEAR = " & sSeason

con.Execute "UPDATE FILE4_10 " & _
            " SET FILE4_10.DESCA = FILE1_10.DESCA " & _
            " FROM FILE4_10 INNER JOIN FILE1_10 ON FILE4_10.MEMBER = FILE1_10.CODE " & _
            " WHERE FILE4_10.YEAR = " & sSeason


'con.CommitTrans
GetFile = True
mylast:
prog1.Value = 0
Exit Function
myerror:
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
GoTo mylast
End Function
Private Sub HideNoPhoto()
Dim i As Long
For i = 1 To grid1.rows - 1
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.RowHidden(i) = chkNoPhoto.Value = 1
    End If
Next
End Sub
