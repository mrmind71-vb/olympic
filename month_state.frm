VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form month_statefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   16125
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   4830
      Begin VB.CommandButton cmd_charge 
         Caption         =   " ÕÊÌ· «·Ì «·Ì «·„’«—Ì›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   135
         Width           =   2580
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«÷«›… «·„ÊŸ›Ì‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   675
      Width           =   1275
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "month_state.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   585
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "month_state.frx":2579
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "month_state.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "month_state.frx":6CFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "month_state.frx":9594
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   4140
         Picture         =   "month_state.frx":BB40
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox xDoc_no_charge 
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
         Left            =   0
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox xDate_trans 
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
         Left            =   0
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   4095
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   2340
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   3300
      Begin VB.CommandButton Command2 
         Height          =   510
         Left            =   1665
         Picture         =   "month_state.frx":E313
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   1590
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ÿ»«⁄… ·ﬂ· „ÊŸ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1590
      End
      Begin VB.CommandButton Command1 
         Caption         =   "√÷«›… »‰›” «·»Ì«‰« "
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   8925
      Begin VB.TextBox xMonth 
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
         Height          =   375
         Left            =   5805
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   510
      End
      Begin VB.TextBox xYear 
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
         Height          =   375
         Left            =   7155
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox xdate 
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
         Height          =   375
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1725
      End
      Begin VB.TextBox xDesca 
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
         Height          =   360
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   7845
      End
      Begin VB.TextBox xdoc_no 
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
         Height          =   375
         Left            =   2610
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "«· «—ÌŒ"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   660
      End
      Begin VB.Label xMonthDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label7 
         Caption         =   "«·»Ì«‰ :"
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
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   675
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "«·‘Â—"
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
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "«·”‰…"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   945
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   1035
      Top             =   1170
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
      Left            =   945
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
      Left            =   90
      Top             =   1215
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
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9045
      Width           =   3165
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   330
         Left            =   45
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame fmTotal 
      Height          =   960
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   8460
      Width           =   6315
      Begin VB.TextBox xTax 
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
         Left            =   3060
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   540
         Width           =   1650
      End
      Begin VB.Label Label2 
         Caption         =   "’«›Ì «·„— »"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   540
         Width           =   1560
      End
      Begin VB.Label Label10 
         Caption         =   "«” ﬁÿ«⁄« "
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   180
         Width           =   1650
      End
      Begin VB.Label Label5 
         Caption         =   "«÷«›« "
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   690
      End
      Begin VB.Label xTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1560
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6630
      Left            =   135
      TabIndex        =   34
      Top             =   1800
      Width           =   15765
      _cx             =   27808
      _cy             =   11695
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
      Rows            =   50
      Cols            =   7
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   8415
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   39
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "month_state.frx":1073D
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "month_state.frx":1290D
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   40
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "month_state.frx":14A55
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "month_state.frx":16C1D
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   41
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "month_state.frx":18D6C
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "month_state.frx":1AF4C
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   42
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "month_state.frx":1D0A7
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "month_state.frx":1F263
      End
   End
End
Attribute VB_Name = "month_statefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[date]", addDate(xdate.Text))
aInsert = AddFlag(aInsert, "[YEAR]", Val(xYear.Text))
aInsert = AddFlag(aInsert, "[MONTH]", Val(xMonth.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
con.BeginTrans
On Error GoTo myerror
If xdoc_no.Text = "" Then
    xdoc_no.Text = Newflag("SALARYH", "doc_no")
    aInsert = AddFlag(aInsert, "[DOC_NO]", addvalue(xdoc_no.Text))
    con.Execute addInsert(aInsert, "salaryH")
Else
    con.Execute addUpdate(aInsert, "SalaryH", "doc_no = " & xdoc_no.Text)
End If
myReplacegrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search.grid1.TextMatrix(Search.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·„ÊŸ› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
        
    grid1.TextMatrix(grid1.Row, 0) = Search.grid1.TextMatrix(Search.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = Search.grid1.TextMatrix(Search.grid1.Row, 1)
    GrdDesc grid1.Row
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.Select grid1.Rows - 1, 1
    End If
    Calctotals
ElseIf ActiveControl.Name = CmdInform.Name Then
    xdoc_no.Text = Search.grid1.TextMatrix(Search.grid1.Row, 0)
    Unload Search
    myUndo
End If
Exit Sub
myerror:
End Sub
Private Sub cmd_charge_Click()
If MsgBox("Â·  Êœ «· ÕÊÌ·", vbOKCancel) <> vbOK Then Exit Sub

xDate_trans.Text = Format(Date, "DD-MM-YYYY")
Set datefrm.oDate = xDate_trans
datefrm.Show 1

Dim cString As String, aCharge As Variant
aCharge = GetRows("Select doc_no from file8_50 where doc_no_salary = " & xdoc_no.Text)
If Not IsEmpty(aCharge) Then
    If MsgBox("«·„·› „ÕÊ· „‰ ﬁ»· Â·  Êœ «· ÕÊÌ· !! ‰⁄„ «„ ·« ?", vbOKCancel) <> vbOK Then Exit Sub
End If

con.BeginTrans
On Error GoTo myerror
If Not IsEmpty(aCharge) Then
    For i = 0 To UBound(aCharge)
        con.Execute "Delete * from file8_50 where doc_no = " & MyParn(retFlag(aCharge(i), "DOC_NO"))
        con.Execute "Delete * from file8_50h where doc_no = " & MyParn(retFlag(aCharge(i), "DOC_NO"))
    Next
End If

Dim loctable As New ADODB.Recordset
cString = "SELECT SALARY.DOC_NO, SALARY.EMPCODE,VAL(SALARY.BASE & '') + VAL(SALARY.MOT & '') +  VAL(SALARY.RW1 & '') +  VAL(SALARY.RW2 & '') +  VAL(SALARY.RW3 & '') +  VAL(SALARY.RW4 & '') - VAL(SALARY.DT & '')-  VAL(SALARY.loan & '')  - VAL(SALARY.INS & '') AS [TOTAL] " & _
          ",EMP.CHARGE,EMP.DESCA FROM SALARY INNER JOIN EMP ON SALARY.EMPCODE = EMP.CODE "
cString = cString & turn(cString) & "SALARY.DOC_NO = " & xdoc_no.Text
cString = cString & turn(cString) & " (NOT EMP.CHARGE IS NULL)"

loctable.Open cString, con, adOpenStatic, adLockReadOnly
Dim aInsert As Variant, sDoc_no As String
Do Until loctable.EOF
    sDoc_no = RetZero(Newflag("FILE8_50H", "DOC_NO"))
    aInsert = AddFlag(Empty, "DOC_NO", addstring(sDoc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate_trans.Text))
    con.Execute addInsert(aInsert, "FILE8_50H")
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(sDoc_no))
    aInsert = AddFlag(aInsert, "BOX", addstring("001"))
    aInsert = AddFlag(aInsert, "CHARGE", addstring(loctable!CHARGE))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!TOTAL & ""))
    aInsert = AddFlag(aInsert, "DOC_NO_SALARY", addvalue(xdoc_no.Text))
    aInsert = AddFlag(aInsert, "DESCA", addstring("„— » " & loctable!Desca & " ⁄‰ ‘Â— " & arbMonth(xMonth.Text)))
    con.Execute addInsert(aInsert, "FILE8_50")
    loctable.MoveNext
Loop
con.CommitTrans
Inform " „  ÕÊÌ· «·„— »«  «·Ì «·„’«—Ì›"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    'on Error GoTo MyError
    con.BeginTrans
    con.Execute "Delete * From salary where Doc_No = " & xdoc_no.Text
    con.Execute "Delete * From salaryH where Doc_No = " & xdoc_no.Text
    con.CommitTrans
    openCardTable
    myUndo
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
'If Not bNoMsgExit Then If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
End Sub
Private Sub cmdSave_Click()
mySave
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdPrint_Click()
    mySave False
    doprint 1
End Sub
Private Sub Command1_Click()
xdoc_no.Text = ""
xdate.Text = Format(Date, "dd-mm-yyyy")
xMonth.Text = ""
xYear.Text = Year(Date)
xDesca.Text = ""
With grid1
For i = 1 To grid1.Rows - 2
    grid1.TextMatrix(i, .Cols - 1) = ""
Next
End With
Handlecontrols DefineMode
End Sub

Private Sub Command3_Click()
mySave False
'doprint App.Path & "\Reports\salary2.rpt"
doprint
End Sub

Private Sub cmdAddEmp_Click()
Dim loctable As New ADODB.Recordset
If Not (IsNumeric(xMonth.Text) And IsNumeric(xYear.Text) And IsDate(xdate.Text)) Then
    MsgBox "«·‘Â— Ê«·”‰… «Ê«· «—ÌŒ €Ì— „”Ã·Ì‰"
    Exit Sub
End If

If Not (Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12) Then
    MsgBox "«·‘Â— €Ì— ’ÕÌÕ"
    Exit Sub
End If

If Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12 Then
    cField1 = "(Select Sum(emploan.value) from emploan where [month] = " & xMonth.Text & _
               " and  [year] = " & xYear.Text & " and empcode =  emp1.code) as loan"
End If

cString = "select code,desca,base,over,ins, " & cField1 & " from emp as emp1 where (isNull(DateEnd) or emp1.DateEnd >= " & DateSq(xdate.Text) & ") order by emp1.group,code"
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
On Error GoTo myerror
Do Until loctable.EOF
    If grid1.FindRow(loctable!Code & "", , 0) = -1 Then
        grid1.TextMatrix(grid1.Rows - 1, 0) = loctable!Code
        grid1.TextMatrix(grid1.Rows - 1, 1) = loctable!Desca
        grid1.TextMatrix(grid1.Rows - 1, 2) = loctable!base & ""
        grid1.TextMatrix(grid1.Rows - 1, 3) = loctable!over & ""
        grid1.TextMatrix(grid1.Rows - 1, 9) = loctable!INS & ""
        grid1.TextMatrix(grid1.Rows - 1, 10) = loctable!Loan & ""
        grid1.AddItem ""
    End If
    loctable.MoveNext
Loop
MsgBox " „ «÷«›… «·„ÊŸ›Ì‰ »‰Ã«Õ"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM SALARYH ORDER BY DOC_NO", con, adOpenKeyset, adLockOptimistic, adCmdText

Set grid1.DataSource = DATA3
DATA3.ConnectionString = con.ConnectionString

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
    Fixgrd
    xdoc_no.Text = ""
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Err.Clear
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 0 Then
    'If Row <> 1 Then grid1.TextMatrix(Row, 1) = myShortCut(Trim(grid1.TextMatrix(Row, 1)), Trim(grid1.TextMatrix(Row - 1, 1)))
    GrdDesc Row
End If
Calctotals
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If grid1.Row <> grid1.Rows - 1 And IsNumeric(grid1.TextMatrix(Row, 0)) And Col = 1 And IsNumeric(xMonth.Text) And IsNumeric(xYear.Text) Then
        cString = "SELECT holidayH.Date,Holiday.Notes " & _
          " FROM HolidayH inner join Holiday on HolidayH.doc_no = Holiday.Doc_no " & _
          " where empcode = " & grid1.TextMatrix(Row, 0) & _
          " and HolidayH.date >= " & DateSq("26-" & IIf(xMonth = "1", "12", Val(xMonth.Text) - 1) & "-" & IIf(Val(xMonth) = 1, Val(xYear.Text) - 1, Val(xYear.Text))) & _
          " and HolidayH.date <= " & DateSq("25-" & Val(xMonth.Text) & "-" & Val(xYear.Text)) & _
          " order by HolidayH.Date"
    holidayDtlfrm.cString = cString
    holidayDtlfrm.Show 1

End If
If grid1.Row <> grid1.Rows - 1 And Trim(grid1.TextMatrix(Row, 0)) <> "" And Col = 8 Then
    doprint 0, grid1.TextMatrix(grid1.Row, 0)
End If
End Sub

Private Sub grid1_EnterCell()
If (grid1.Col = 11) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.Select 1, 1
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then
    grid1.AddItem "", grid1.Row
End If
If KeyCode = 112 Then
    If grid1.Col = 0 And grid1.Row <> 0 Then empLookup Me, Search
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
End If
End Sub
Private Sub xDiscount_LostFocus()
Calctotals
End Sub
Private Function MYVALID() As Boolean

'If xdoc_no.Text = "" Then
'    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
'    Exit Function
'End If

If Val(xYear.Text) < 2010 Or Val(xYear.Text) > 2030 Or Not IsNumeric(xYear.Text) Then
    MsgBox "«·”‰… €Ì— ”·Ì„…"
    Exit Function
End If
If Val(xMonth) < 1 Or Val(xMonth) > 12 Then
    MsgBox "«·‘Â— €Ì— ”·Ì„"
    Exit Function
End If
If Not IsDate(xdate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

With grid1
For i = 1 To grid1.Rows - 2
    If Not validRow(i) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xdoc_no.Text = CardTable!doc_no
xdate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xMonth.Text = CardTable!Month & ""
xMonthDesca.Caption = arbMonth(Val(CardTable!Month & ""))
xYear.Text = CardTable!Year & ""
xDesca.Text = CardTable!Desca & ""

xDesca_w1.Text = CardTable!Desca_w1 & ""
xDesca_w2.Text = CardTable!Desca_w2 & ""
xDesca_w3.Text = CardTable!Desca_w3 & ""
xDesca_w4.Text = CardTable!Desca_w4 & ""
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub mydefine()
xdoc_no.Text = ""
xdate.Text = Format(Date, "dd-mm-yyyy")
xMonthDesca.Caption = ""
xMonth.Text = ""
xYear.Text = Year(Date)
xDesca.Text = ""

xTotal1.Caption = ""
xTotal2.Caption = ""
xtotal.Caption = ""

xDesca_w1.Text = ""
xDesca_w2.Text = ""
xDesca_w3.Text = ""
xDesca_w4.Text = ""
grid1.Rows = 1
grid1.AddItem ""
Fixgrd
'grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdPrint.Enabled = (nMode = LoadMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdSave.Enabled = (bedit) And (Not isOut) Or nMode = DefineMode
CmdDelInv.Enabled = nMode = LoadMode And bedit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode)
End Sub

Private Sub XDATE_DblClick()
Set datefrm.oDate = xdate
datefrm.Show 1
End Sub

Private Sub xDoc_No_LostFocus()
If xdoc_no.Text = "" Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_no = " & MyParn(xdoc_no.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "Delete * from  salary where autocode = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
         grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 0
        If KeyCode = 112 Then
            empLookup Me, Search
        End If
End Select
End Sub
Private Sub GrdDesc(Row)
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
grid1.TextMatrix(Row, 1) = GetDesca("Select desca from emp where code = " & grid1.TextMatrix(Row, 0))
End Sub
Private Function Calctotals()
Dim nTotal As Single, nTotal2 As Single, nTotal3 As Single
With grid1
For i = 1 To grid1.Rows - 2
    grid1.TextMatrix(i, 8) = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7))
    nTotal1 = nTotal1 + Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7))
    grid1.TextMatrix(i, 12) = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 10)) - Val(.TextMatrix(i, 11))
    nTotal2 = nTotal2 + Val(.TextMatrix(i, 9)) + Val(.TextMatrix(i, 10)) + Val(.TextMatrix(i, 11))
    nTotal = nTotal + Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4)) + Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 9)) - Val(.TextMatrix(i, 10)) - Val(.TextMatrix(i, 11))
Next
xTotal1.Caption = Format(nTotal1, "Fixed")
xTotal2.Caption = Format(nTotal2, "Fixed")
xtotal.Caption = Format(nTotal, "Fixed")
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO, [Month] & '-' & [Year],Format([DATE],'yyyy/mm/dd'),Desca " & _
                  " FROM  SALARYH"

Generalarray(2) = "Order by Year,Month,DOC_NO "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·‘Â—-«·”‰…-«· «—ÌŒ-«·Ê’›"
listarray(0, 1) = "( val('cFilter') =  [year] or val('cFilter') =  [Month] or %%Desca%% " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·‘Â—-«·”‰…"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1200

GrdArray(3, 0) = "»Ì«‰"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search
Search.Caption = "«” ⁄·«„"
Search.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow And Trim(grid1.TextMatrix(i, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·„ÊŸ› " & grid1.TextMatrix(nRow, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function
Private Function validRow(nRow) As Boolean
If nRow > 0 Then
    If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Function
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
End If
validRow = True
End Function
Private Sub Fixgrd()
With grid1
.MergeCells = flexMergeFree
.MergeRow(0) = True
.Cell(flexcpAlignment, 0, 4, 0, 7) = flexAlignCenterCenter
.FormatString = "ﬂÊœ|" & "«·«”„|" & "«·—« » «·«”«”Ì|" & "«·„ €Ì—|" & "„ﬂ«›√… √Œ—Ì|" & "„ﬂ«›√… √Œ—Ì|" & "„ﬂ«›√… √Œ—Ì|" & "„ﬂ«›√… √Œ—Ì|" & "«·«Ã„«·Ì|" & "«· √„Ì‰« |" & "«·«” ﬁÿ«⁄« |" & "”·›|" & "«·’«›Ì|" & "„·«ÕŸ« |"
.ColWidth(0) = 600
.ColWidth(1) = 2800
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 800
.ColWidth(5) = 800
.ColWidth(6) = 800
.ColWidth(7) = 800
.ColWidth(8) = 1000
.ColWidth(9) = 800
.ColWidth(10) = 800
.ColWidth(11) = 800
.ColWidth(12) = 800
.ColWidth(13) = 3000

.ColComboList(1) = "..."
.ColComboList(8) = "..."

.ColHidden(.Cols - 1) = True
.ColFormat(3) = "#.##"
.ColFormat(4) = "#.##"
.ColFormat(5) = "#.##"
.ColFormat(6) = "#.##"
.ColFormat(7) = "#.##"
.ColFormat(8) = "#.##"
.ColFormat(9) = "#.##"
.ColFormat(10) = "#.##"
.ColFormat(11) = "#.##"
.ColFormat(12) = "#.##"

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub doprint(Optional nFlag As Long = 0, Optional cEmpCode As String = "")
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    If cEmpCode = "" Or (Trim(grid1.TextMatrix(i, 0)) = cEmpCode) Then
        temptable.AddNew
        temptable!str21 = "ﬂ‘› „— »«  «·⁄«„·Ì‰ »«·„Ã·… ( Õ—Ì—) ⁄‰ ‘Â— " & xMonthDesca.Caption & Space(1) & "”‰… " & xYear.Text
        temptable!str1 = TurnValue(grid1.TextMatrix(i, 1))
        
        temptable!val1 = Val(grid1.TextMatrix(i, 2))
        temptable!val2 = Val(grid1.TextMatrix(i, 3))
        
        temptable!Val3 = Val(grid1.TextMatrix(i, 4))
        
        temptable!val5 = Val(grid1.TextMatrix(i, 5))
        temptable!Val6 = Val(grid1.TextMatrix(i, 6))
        temptable!Val7 = Val(grid1.TextMatrix(i, 7))
        
        temptable!Val8 = Val(grid1.TextMatrix(i, 8))
        temptable!val9 = Val(grid1.TextMatrix(i, 9))
        temptable!Val10 = Val(grid1.TextMatrix(i, 10))
        temptable!val11 = Val(grid1.TextMatrix(i, 11))
        temptable!val12 = Val(grid1.TextMatrix(i, 12))
        temptable!str2 = TurnValue(grid1.TextMatrix(i, 13))
         
        temptable!Str3 = TurnValue(xDesca_w1.Text)
        temptable!str4 = TurnValue(xDesca_w2.Text)
        temptable!str5 = TurnValue(xDesca_w3.Text)
        temptable!str6 = TurnValue(xDesca_w4.Text)
        temptable!Val20 = nFlag
        
        temptable.Update
    End If
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\salary2.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub mySave(Optional bMsg As Boolean = True)
If Not MYVALID Then Exit Sub
Calctotals
If Not myreplace Then Exit Sub
CardTable.Requery
If bMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Doc_No = " & xdoc_no.Text, , adSearchForward, adBookmarkFirst
'Handlecontrols LoadMode
'xBalance.Caption = Format(GetDesca("Select sum(val(SAL & '') - val (pay & '')) as balance FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Sub myproc2(nDoc_no)
bNoMsgExit = True
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Private Sub myReplacegrd()
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
With grid1
    For i = 1 To .Rows - 2
        prog1.Value = Round(i / (grid1.Rows - 2), 2) * 100
        aInsert = AddFlag(Empty, "[DOC_NO]", addvalue(xdoc_no.Text))
        aInsert = AddFlag(aInsert, "[EMPCODE]", addvalue(.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[BASE]", Val(.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[MOT]", Val(.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "[RW1]", Val(.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "[RW2]", Val(.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "[RW3]", Val(.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "[RW4]", Val(.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "[INS]", Val(.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "[DT]", Val(.TextMatrix(i, 10)))
        aInsert = AddFlag(aInsert, "[LOAN]", Val(.TextMatrix(i, 11)))
        aInsert = AddFlag(aInsert, "[NOTES]", addstring(.TextMatrix(i, 13)))
        aInsert = AddFlag(aInsert, "[ROW]", i)
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "salary")
        Else
            con.Execute addUpdate(aInsert, "salary", "autocode = " & .TextMatrix(i, .Cols - 1))
        End If
        prog1.Visible = False
    Next
End With
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE2_50., DRIVER.DESCA, FILE2_50.DAYS," & _
              " FROM FILE2_10 inner JOIN EMP ON SALARY.EMPCODE = EMP.CODE WHERE DOC_NO = " & xdoc_no.Text & " order by SALARY.ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
    grid1.AddItem ""
End With
Calctotals
Fixgrd
End Sub
Private Sub xMonth_Validate(Cancel As Boolean)
If Trim(xMonth.Text) = "" Then Exit Sub
If IsNumeric(xMonth.Text) Then
    If xMonth.Text < 1 Or xMonth.Text > 12 Then
        MsgBox "«·‘Â— ·« Ì’·Õ"
        Cancel = True
        Exit Sub
    End If
    xMonthDesca.Caption = arbMonth(xMonth.Text)
    If IsNumeric(xMonth.Text) Then
        If Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12 Then
            Dim sDoc_no As Variant
            sDoc_no = GetField("Select doc_no from salaryh where [YEAR] = " & xYear.Text & _
                " and [Month] = " & xMonth.Text)
            If Not IsEmpty(sDoc_no) Then
                xdoc_no.Text = sDoc_no
                myUndo
            Else
                mydefine
            End If
        End If
    End If
End If
End Sub
Private Sub xYear_Validate(Cancel As Boolean)
If Trim(xYear.Text) = "" Then Exit Sub
If IsNumeric(xYear.Text) Then
    If Val(xYear.Text) < 2010 Or Val(xYear.Text) > 2030 Then
        MsgBox "«·”‰… ·«  ’·Õ"
        Cancel = True
        Exit Sub
    End If
    If IsNumeric(xMonth.Text) Then
        If Val(xMonth.Text) >= 1 And Val(xMonth.Text) <= 12 Then
            Dim sDoc_no As Variant
            sDoc_no = GetField("Select doc_no from salaryh where [YEAR] = " & xYear.Text & _
                " and [Month] = " & xMonth.Text)
            If Not IsEmpty(sDoc_no) Then
                xdoc_no.Text = sDoc_no
                myUndo
            Else
                mydefine
            End If
        End If
    End If
End If
End Sub

