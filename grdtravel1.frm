VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form grdTravelfrm1 
   Caption         =   "„ «»⁄… «·⁄„·Ì« "
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
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   675
      Width           =   9285
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’œ— ·Â« ›« Ê—…"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   135
         Width           =   1725
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’œ— ·Â« »Ê·Ì’… Ê·„ Ì’œ— ·Â« ›« Ê—…"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   135
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "·„ Ì’œ— ·Â« »Ê·Ì’…"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   135
         Width           =   1860
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ·"
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
         Left            =   8505
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   135
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1305
      Width           =   5595
      Begin VB.CommandButton cmdClear 
         Height          =   555
         Left            =   1140
         Picture         =   "grdtravel1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   2235
         Picture         =   "grdtravel1.frx":2424
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   4455
         Picture         =   "grdtravel1.frx":4C0F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "grdtravel1.frx":7101
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   3330
         Picture         =   "grdtravel1.frx":956D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2040
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
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox xDate_Policy2 
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
         TabIndex        =   7
         Tag             =   "D"
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDate_Policy1 
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "D"
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xPolicy 
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
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   3030
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
         Left            =   8685
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "new2"
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox xCar 
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
         MaxLength       =   5
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
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
         TabIndex        =   2
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
         Left            =   90
         TabIndex        =   9
         Tag             =   "new"
         Top             =   900
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   90
         TabIndex        =   10
         Tag             =   "new"
         Top             =   1260
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSDataListLib.DataCombo xCargo 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Tag             =   "new"
         Top             =   1620
         Width           =   2310
         _ExtentX        =   4075
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
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·Õ„Ê·…"
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
         Index           =   4
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1665
         Width           =   1155
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
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1260
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
         TabIndex        =   39
         Tag             =   "t"
         Top             =   1260
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1305
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·»Ê·Ì’…"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label16 
         Caption         =   "—ﬁ„ «·»Ê·Ì’…"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   540
         Width           =   1110
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label6 
         Caption         =   "Õ Ï"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   495
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
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Tag             =   "t"
         Top             =   1620
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
         Left            =   9855
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1620
         Width           =   510
      End
      Begin VB.Label xCar_Desca 
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
         TabIndex        =   25
         Tag             =   "t"
         Top             =   900
         Width           =   2265
      End
      Begin VB.Label xCar_type 
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "t"
         Top             =   900
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "«·”Ì«—…"
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
         Left            =   9900
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   945
         Width           =   555
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
      Left            =   90
      TabIndex        =   21
      Top             =   2070
      Width           =   20085
      _cx             =   35428
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
      Cols            =   14
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
      TabIndex        =   22
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
Attribute VB_Name = "grdTravelfrm1"
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
Private Sub Cmd_Print_Click()
    
End Sub

Private Sub cmdDelinv_Click()
End Sub

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
End Sub
Private Sub cmdPrint_Click()
Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
Dim aHeader As Variant
cHeader1 = "„ «»⁄… «Ã„«·Ì ”›—Ì«  Œ·«· › —…"
If IsDate(xDate1.Text) Or IsDate(xdate2.Text) Then aHeader = AddFlag(aHeader, BetweenString(Format(xDate1.Text, "YYYY-MM-DD"), xdate2.Text))
If IsDate(xDate_Policy1.Text) Or IsDate(xDate_Policy2.Text) Then aHeader = AddFlag(aHeader, BetweenString(xDate_Policy1.Text, xDate_Policy2.Text, "„‰  «—ÌŒ »Ê·Ì’…"))
If Trim(xcar.Text) <> "" Then aHeader = AddFlag(aHeader, "”Ì«—… —ﬁ„ : " & xcar.Text & turn(xCar_Desca.Caption, " ") & xCar_Desca.Caption & turn(xCar_type.Caption, " ") & xCar_type.Caption)
If Trim(xCode_sup.Text) <> "" Then aHeader = AddFlag(aHeader, "«·„Ê—œ : " & Me.xCode_sup_desca.Caption)
If Trim(xCode.Text) <> "" Then aHeader = AddFlag(aHeader, "«·⁄„Ì· : " & xcode_Desca.Caption)
If Trim(xPlace1.Text) <> "" Then aHeader = AddFlag(aHeader, "„‰ : " & xPlace1.Text)
If Trim(xPlace2.Text) <> "" Then aHeader = AddFlag(aHeader, "Õ Ï : " & xPlace2.Text)
If Trim(xPolicy.Text) <> "" Then aHeader = AddFlag(aHeader, "—ﬁ„ «·»Ê·Ì’… : " & xPolicy.Text)
If xDriver.MatchedWithList Then aHeader = AddFlag(aHeader, "«·”«∆ﬁ : " & xDriver.Text)
If Not IsEmpty(aHeader) Then
    cHeader2 = retHeader(aHeader, 0, 2)
    cHeader3 = retHeader(aHeader, 2, 2)
    cHeader4 = retHeader(aHeader, 4, 2)
End If
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 11)
PrintGrdNew.doprint grid1, 0.84, -2, cHeader1, cHeader2, cHeader3, , False, True, 9, , aRow
PrintGrdNew.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
openCon con
Set data1.Recordset = myRecordSet("Select Code,DescA From driver where driver = 1 order by Desca", con)
Set xDriver.RowSource = data1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"
    
Set DATA2.Recordset = myRecordSet("SELECT * FROM PLACE_CODES ORDER BY DESCA", con)
Set xPlace1.RowSource = DATA2
xPlace1.ListField = "DESCA"
xPlace1.BoundColumn = "CODE"
    
Set xPlace2.RowSource = DATA2
xPlace2.ListField = "DESCA"
xPlace2.BoundColumn = "CODE"
        
Set DATA3.Recordset = myRecordSet("SELECT * FROM CARGO_CODES", con)
Set xCargo.RowSource = DATA3
xCargo.ListField = "Desca"
xCargo.BoundColumn = "Code"

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon
Fixgrd
grid1.Rows = 1
LoadText Me
'GetCaption
End Sub
Private Sub myload()
Dim cString As String
cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY,Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),DRIVER.DESCA + CASE WHEN DRIVER_B.DESCA IS NULL THEN '' ELSE '-' + DRIVER_B.DESCA END   ,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL,FILE6_20.DOC_NO,CARGO_CODES.DESCA  " & _
        " FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE LEFT JOIN FILE6_20 ON TRAVEL_H.DOC_NO = FILE6_20.TRAVEL" & _
        " LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE " & _
        " LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
        " LEFT JOIN DRIVER AS DRIVER_B ON TRAVEL_H.DRIVER2 = DRIVER_B.CODE" & _
        " LEFT JOIN FILE4_10 ON TRAVEL_H.CODE_SUP = FILE4_10.CODE" & _
        " LEFT JOIN TRAVEL_C ON TRAVEL_H.DOC_NO = TRAVEL_C.DOC_NO" & _
        " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
        " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE" & _
        " LEFT JOIN CARGO_CODES ON TRAVEL_H.CARGO = CARGO_CODES.CODE"

'cString = "SELECT TRAVEL_H.DOC_NO,Convert(VARCHAR(10),TRAVEL_H.[DATE],111),FILE3_10.Desca,TRAVEL_H.POLICY,Convert(VARCHAR(10),TRAVEL_H.[DATE_POLICY],111),DRIVER.DESCA + CASE WHEN DRIVER_B.DESCA IS NULL THEN '' ELSE '-' + DRIVER_B.DESCA END   ,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL,CASE WHEN CODE_SUP IS NULL THEN  SUM(COALESCE(TRAVEL_C.[VALUE],0)) ELSE TOTAL_SUP END,TRAVEL_H.TOTAL - CASE WHEN CODE_SUP IS NULL THEN SUM(COALESCE(TRAVEL_C.[VALUE],0)) ELSE TOTAL_SUP  END,FILE6_20.DOC_NO " & _
'        " FROM  TRAVEL_H INNER JOIN FILE3_10 ON TRAVEL_H.CODE  = FILE3_10.CODE LEFT JOIN FILE6_20 ON TRAVEL_H.DOC_NO = FILE6_20.TRAVEL" & _
'        " LEFT JOIN CARS ON TRAVEL_H.CAR = CARS.CODE " & _
'        " LEFT JOIN DRIVER ON TRAVEL_H.DRIVER = DRIVER.CODE" & _
'        " LEFT JOIN DRIVER AS DRIVER_B ON TRAVEL_H.DRIVER2 = DRIVER_B.CODE" & _
'        " LEFT JOIN FILE4_10 ON TRAVEL_H.CODE_SUP = FILE4_10.CODE" & _
'        " LEFT JOIN TRAVEL_C ON TRAVEL_H.DOC_NO = TRAVEL_C.DOC_NO" & _
'        " LEFT JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE " & _
'        " LEFT JOIN PLACE_CODES AS PLACE_CODES_B ON TRAVEL_H.PLACE2 = PLACE_CODES_B.CODE"
        
If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE >= " & DateSq(xDate1.Text)
End If

If IsDate(xdate2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE <= " & DateSq(xdate2.Text)
End If

If IsDate(xDate_Policy1.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY >= " & DateSq(xDate_Policy1.Text)
End If

If IsDate(xDate_Policy2.Text) Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY <= " & DateSq(xDate_Policy2.Text)
End If

If Trim(xcar.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.CAR = " & MyParn(xcar.Text)
End If

If Trim(xCode.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.CODE = " & MyParn(xCode.Text)
End If

If Trim(xCode_sup.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.CODE_SUP = " & MyParn(xCode_sup.Text)
End If

If Trim(xPolicy.Text) <> "" Then
    cString = cString & turn(cString) & "TRAVEL_H.POLICY = " & MyParn(xPolicy.Text)
End If

If xPlace1.MatchedWithList Then
    cString = cString & turn(cString) & "PLACE1 = " & xPlace1.BoundText
End If

If xPlace2.MatchedWithList Then
    cString = cString & turn(cString) & "PLACE2 = " & xPlace2.BoundText
End If

If xCargo.MatchedWithList Then
    cString = cString & turn(cString) & "CARGO = " & xCargo.BoundText
End If


If xDriver.MatchedWithList Then
    cString = cString & turn(cString) & "(TRAVEL_H.DRIVER = " & MyParn(xDriver.BoundText) & " OR TRAVEL_H.DRIVER2 = " & MyParn(xDriver.BoundText) & ")"
End If

If Option2.Value Then
    cString = cString & turn(cString) & "TRAVEL_H.DATE_POLICY IS NULL"
ElseIf Option3.Value Then
    cString = cString & turn(cString) & "(NOT TRAVEL_H.DATE_POLICY IS NULL) AND (FILE6_20.TRAVEL Is Null)"
ElseIf Option4.Value Then
    cString = cString & turn(cString) & "(NOT FILE6_20.TRAVEL Is Null)"
End If

cString = cString & " GROUP BY  TRAVEL_H.DOC_NO,TRAVEL_H.[DATE],FILE3_10.Desca,TRAVEL_H.POLICY,TRAVEL_H.[DATE_POLICY],DRIVER.DESCA,DRIVER_B.DESCA,CARS.BOARD,FILE4_10.DESCA,PLACE_CODES.DESCA,PLACE_CODES_B.DESCA,TRAVEL_H.TOTAL,TRAVEL_H.CODE_SUP,TRAVEL_H.TOTAL_SUP,FILE6_20.DOC_NO,CARGO_CODES.DESCA"
Set DATA11.Recordset = myRecordSet(cString, con)
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
    .TextMatrix(0, 12) = "«·›« Ê—…"
    .TextMatrix(0, 13) = "‰Ê⁄ «·Õ„Ê·…"
    
    '.FrozenCols = 2
    .ColWidth(0) = 700
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
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    '.ColHidden(12) = True
    '.ColHidden(13) = True
    
    .ColDataType(11) = flexDTDouble
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    For i = 1 To grid1.Rows - 1
        grid1.TextMatrix(i, 0) = i
    Next

    .ExplorerBar = flexExSortShow
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 12, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 13, "#0", vbRed, vbYellow, True, "  "
    StatusBar1.Panels(1).Text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
    If .Rows > 1 Then
        For i = 0 To 10
            .TextMatrix(1, i) = "«·≈Ã„«·Ì"
        Next
        .MergeRow(1) = True
    End If
    .MergeCells = flexMergeFree

    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Unload Me
Set grditem1 = Nothing
End Sub

Private Sub Grid1_DblClick()
If grid1.Row < 2 Then Exit Sub
If grid1.Row > 0 Then
    Travelfrm.sDoc_no = grid1.TextMatrix(grid1.Row, 1)
    Travelfrm.Show
End If
End Sub
Sub myProc()
If ActiveControl.Name = xcar.Name Then
    xcar.Text = oSearchCar.grid1.TextMatrix(oSearchCar.grid1.Row, 0)
    xCar_Validate False
    Unload oSearchCar
ElseIf ActiveControl.Name = xCode_sup.Name Then
    xCode_sup.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    xCode_sup_Validate False
    Unload oSearchSup
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    xCode_Validate False
    Unload oSearchClient
End If
End Sub
Private Sub grid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'ItemsLookupAll Me, osearchitem
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
With grid2
If grid2.Row = grid2.Rows - 1 Then
    myAddItem
End If
End With
End Sub
Private Sub grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If (Not validRow(OldRow, True, True)) And OldRow <> grid2.Rows - 1 And OldRow <> 0 And grid2.TextMatrix(OldRow, grid2.Cols - 1) = "" Then
    grid2.RemoveItem OldRow
    grid2.SaveGrid cFilesave, flexFileData
End If
End Sub
Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 Then
    grid2.RemoveItem grid2.Row
    grid2.SaveGrid cFilesave, flexFileData
    grid2.Select grid2.Rows - 1, 1
    grid2.ShowCell grid2.Rows - 1, 1
End If
End Sub

Private Sub grid2_Validate(Cancel As Boolean)
If (Not validRow(grid2.Row, True, True)) And grid2.Row <> grid2.Rows - 1 And grid2.Row <> 0 And grid2.TextMatrix(grid2.Row, grid2.Cols - 1) = "" Then
    grid2.RemoveItem grid2.Row
    grid2.SaveGrid cFilesave, flexFileData
End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub myAddItem()
grid2.AddItem ""
grid2.ShowCell grid2.Rows - 1, 1
grid2.SaveGrid cFilesave, flexFileData
End Sub
Private Sub xCar_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    carsLookupAll Me, oSearchCar
End If
End Sub
Private Sub xCar_Validate(Cancel As Boolean)
xCar_Desca.Caption = ""
xCar_type.Caption = ""
If Not ValidInt(xcar.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select code,[TYPE],MODEL,BOARD,gas1 from cars where code = " & xcar.Text)
If IsEmpty(aRet) Then
    MsgBox "ﬂÊœ «·”Ì«—… €Ì— ’ÕÌÕ"
    Cancel = True
Else
    xCar_Desca.Caption = retFlag(aRet, "Board")
    xCar_type.Caption = retFlag(aRet, "TYPE") & " " & retFlag(aRet, "Model")
End If
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
Private Sub GetCaption()
'xCar_Desca.Caption = ""
'xCar_type.Caption = ""
'If Not ValidInt(xCar.Text) Then Exit Sub
'Dim aRet As Variant
'aRet = GetFields("select code,[TYPE],MODEL,BOARD,gas1 from cars where code = " & xCar.Text)
'xCar_Desca.Caption = retFlag(aRet, "Board")
'xCar_type.Caption = retFlag(aRet, "TYPE") & " " & retFlag(aRet, "Model")
'If Not xCode_sup.Text = "" Then Exit Sub
End Sub
Private Sub xDate_Policy2_GotFocus()
myGotFocus xDate_Policy2
End Sub
Private Sub xDate_Policy2_LostFocus()
myLostFocus xDate_Policy2
myValidDate xDate_Policy2
End Sub
Private Sub xDate_policy1_GotFocus()
myGotFocus xDate_Policy1
End Sub
Private Sub xDate_policy1_LostFocus()
myLostFocus xDate_Policy1
myValidDate xDate_Policy1
End Sub

Private Sub xPolicy_GotFocus()
myGotFocus xPolicy
End Sub
Private Sub xPolicy_LostFocus()
myLostFocus xPolicy
End Sub
Private Sub xCode_sup_GotFocus()
myGotFocus xCode_sup
End Sub
Private Sub xCode_sup_LostFocus()
myLostFocus xCode_sup
End Sub
Private Sub xCar_GotFocus()
myGotFocus xcar
End Sub
Private Sub xCar_LostFocus()
myLostFocus xcar
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
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
