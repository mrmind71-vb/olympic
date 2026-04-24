VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form impcostfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬂ·›… «” Ì—«œÌ…"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   15195
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   2475
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   0
      Width           =   6540
      Begin VB.CommandButton cmduntrans 
         Caption         =   "«·€«¡ «· ÕÊÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   135
         Width           =   1635
      End
      Begin VB.CommandButton CMDTRANS 
         Caption         =   " ÕÊÌ· ··„‘ —Ì« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1755
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   135
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«· ”⁄Ì—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3420
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   135
         Width           =   1410
      End
      Begin VB.CommandButton cmdCalcCost 
         Caption         =   "⁄—÷ «· ﬂ·›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4860
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   135
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmditem 
      Caption         =   " ⁄œÌ· ’‰›"
      Height          =   465
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   9465
      Width           =   1995
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   -3555
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2430
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   150
         Width           =   3510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9045
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmd_print 
         Caption         =   "ÿ»«⁄…"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
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
         Left            =   3870
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
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
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   1845
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   585
      Width           =   13335
      Begin VB.CommandButton Command3 
         Caption         =   " Ã„Ì⁄ ›Ï „” ‰œ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1530
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.TextBox xDateTransPur 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   900
         Width           =   1290
      End
      Begin VB.TextBox xDateTrans 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   540
         Width           =   1290
      End
      Begin VB.TextBox xcurRate 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox xFactName 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1620
         Width           =   4395
      End
      Begin VB.TextBox xPolicy 
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
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1260
         Width           =   2445
      End
      Begin VB.TextBox xBankName 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1260
         Width           =   4395
      End
      Begin VB.CheckBox xTrans 
         Alignment       =   1  'Right Justify
         Caption         =   " —ÕÌ· «·Ì «·„Œ“‰"
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
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox xCredit 
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
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   900
         Width           =   2445
      End
      Begin VB.TextBox xVessel 
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
         Height          =   315
         Left            =   7695
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   900
         Width           =   4395
      End
      Begin VB.TextBox xCode 
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
         Height          =   315
         Left            =   11025
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1065
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   315
         Left            =   9855
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2235
      End
      Begin VB.TextBox xDate 
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
         Left            =   4050
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   4050
         TabIndex        =   3
         Top             =   540
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcurrency 
         Height          =   315
         Left            =   4050
         TabIndex        =   43
         Top             =   1620
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ÕÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ÕÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1305
         Width           =   45
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„⁄«„· «· ÕÊÌ· :"
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„·… :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1695
         Width           =   540
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê’› «·—”«·…"
         Height          =   195
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1665
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»Ê·Ì’… :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "«”„ «·»‰ﬂ :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·«⁄ „«œ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label7 
         Caption         =   "≈”„ «·„—ﬂ» :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   990
         Width           =   1020
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   3285
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         Height          =   240
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   255
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
         Height          =   195
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
         Height          =   195
         Left            =   12150
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1530
      Width           =   1635
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   1455
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2070
      Top             =   -90
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
      Height          =   5895
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2610
      Width           =   14970
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5685
         Left            =   90
         TabIndex        =   62
         Top             =   180
         Width           =   14820
         _cx             =   26141
         _cy             =   10028
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
         Rows            =   50
         Cols            =   10
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
   End
   Begin VB.Frame Frame7 
      Height          =   1455
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8505
      Width           =   11010
      Begin VB.CommandButton cmdCharge 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   900
         Width           =   330
      End
      Begin VB.TextBox xDiscount 
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
         Left            =   7650
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label xTotalNoCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "»«·⁄„·… «·„Õ·Ì… :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "»⁄œ «·Œ’„ :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
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
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label xTotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label xTotalFrgn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label xCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì «· ﬂ·›… :"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "„’«—Ì› :"
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
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "«·Œ’„ :"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   690
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label9 
         Caption         =   "≈Ã„«·Ì «·√’‰«› :"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   630
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -2115
      Top             =   990
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
   Begin VB.Frame Frame10 
      Height          =   930
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8505
      Width           =   3930
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   74
         ToolTipText     =   "»ÕÀ"
         Top             =   525
         Width           =   2925
      End
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   5
         ToolTipText     =   "»ÕÀ"
         Top             =   135
         Width           =   2925
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ’‰›"
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
         Height          =   240
         Left            =   3075
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ﬂÊœ"
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
         Height          =   240
         Left            =   3075
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   225
         Width           =   780
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame9 
      Height          =   570
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   9375
      Width           =   1920
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -1485
      Top             =   495
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   180
      TabIndex        =   63
      Top             =   9960
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
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
End
Attribute VB_Name = "impcostfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim tBalance  As ADODB.Recordset
Dim cFile As String, cFileClient, cMoveName, cFileMove, cItemmove As String, cClientmove, cFieldItem, cFieldClient
Dim formMode, dDateLast As String
Dim cList As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
If Not MYVALID Then Exit Function
Dim nTry As Integer, nAffected As Integer, aInsert(13, 1), aGrid(13, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "code"
aInsert(2, 1) = addstring(XCODE.Text)

aInsert(3, 0) = "Datetranspur"
aInsert(3, 1) = addDate(xDateTransPur.Text)

aInsert(4, 0) = "Credit"
aInsert(4, 1) = addstring(xCredit.Text)

aInsert(5, 0) = "Vessel"
aInsert(5, 1) = addstring(xVessel.Text)

aInsert(6, 0) = "BankName"
aInsert(6, 1) = addstring(xBankName.Text)

aInsert(7, 0) = "Policy"
aInsert(7, 1) = addstring(xPolicy.Text)

aInsert(8, 0) = "FactName"
aInsert(8, 1) = addstring(XFACTNAME.Text)

aInsert(9, 0) = "[Currency]"
aInsert(9, 1) = addstring(xcurrency.BoundText)

aInsert(10, 0) = "CurRate"
aInsert(10, 1) = Val(xcurRate.Text)

aInsert(11, 0) = "Trans"
aInsert(11, 1) = IIf(xtrans.Value = 0, "0", "1")

aInsert(12, 0) = "Discount"
aInsert(12, 1) = Val(xDiscount.Text)

aInsert(13, 0) = "Charge"
aInsert(13, 1) = Val(xcharge.Caption)

con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE7_60H", "doc_no")), 15)
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "File7_60H")
Else
    con.Execute CreateUpdate(aInsert, "FILE7_60H", " WHERE FILE7_60H.DOC_NO = " & MyParn(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    GrdDesc grid1.Row
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = 1
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
        MakeSerial
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 3) = 1
        grid1.Select grid1.Rows - 1, 1
    End If
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    Calctotals
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    Unload Search3
    myload
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
Exit Sub
myerror:
Unload Search
End Sub
Private Sub cmdaddGroup_Click()
ReDim aPublic(0)
Set aPublic(0) = Me
additemfrm.Show 1
End Sub
Private Sub cmdCopy_Click()
If Not MYVALID Then Exit Sub
myreplace
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
copyPurFrm.Show 1
End Sub

Private Sub cmdCalcCost_Click()
'If Not MYVALID Then Exit Sub
'CalcCost
'If MyReplace <> 0 Then
'    MsgBox "·„ Ì „ Õ”«» «· ﬂ·›… ·ÊÃÊœ „‘ﬂ·… ›Ï «·Õ›Ÿ"
'    Exit Sub
'End If
'MsgBox "ÌÃ» Õ›Ÿ «·»Ì«‰«  «–« ÕœÀ  ⁄œÌ· ·÷„«‰ œﬁ… Õ”«»… «· ﬂ·›… «·«” Ì—«œÌ…"
'CalcCost
'CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
mysave True
doprint
If MsgBox(" ÕœÌÀ  ﬂ·›… «·√’‰«› «·Œ«’… »«·›« Ê—… «·≈” Ì—«œÌ…", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE1_10 SET FILE1_10.COST = [FILE7_60].[COST] FROM FILE1_10 LEFT JOIN FILE7_60 ON FILE1_10.ITEM = FILE7_60.ITEM WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text)
    MsgBox " „  ÕœÌÀ"
End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
   ' Õ–› «·„” ‰œ
    con.Execute "Delete  From FILE7_60 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE7_60H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.BOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
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

Private Sub cmdLookup2_Click()
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , CONVERT(VARCHAR(10),[DATE],111), FILE4_10.Desca,Credit,Vessel" & _
                  " FROM  (FILE7_60H left JOIN FILE4_10 ON FILE7_60H.CODE = FILE4_10.CODE )"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ-«”„ «·”›Ì‰…"
listarray(0, 1) = "(@@Doc_No@@ or  %%FILE4_10.DESCA%% OR %%VESSEL%% OR" & _
                  " ##[Date]##)"

listarray(1, 0) = "—ﬁ„ «·«⁄ „«œ"
listarray(1, 1) = "(Credit Like '%cFilter%')"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ «·«⁄ „«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«”„ «·„—ﬂ»"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "≈Ã„«·Ì «·ﬂ„Ì…"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
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
'xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
    mysave
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.BOF And CardTable.EOF Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub cmdItem_Click()
Dim bEditLocal As Boolean
bEditLocal = bEdit: bEdit = True
itemsfrm.Show
bEdit = bEditLocal
End Sub

Private Sub cmdCharge_Click()
impchargefrm.Show 1
xcharge.Caption = GetDesca("Select sum([value]) from fiLE7_60CH WHERE DOC_NO = " & MyParn(xDoc_No.Text))
Calctotals
End Sub
Private Sub cmduntrans_Click()
    If myunTrans Then
        CmdUndo_Click
        MsgBox " „ «·€«¡ «· —ÕÌ· »‰Ã«Õ"
    End If
End Sub

Private Sub Command1_Click()
If grid1.Rows <= 2 Then Exit Sub
impcostpricefrm.Show 1
myload
End Sub
Private Sub Cmd_Print_Click()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str21 = "ÿ»«⁄… „” ‰œ ≈” Ì—«œ "
    temptable!str1 = xDoc_No.Text
    temptable!str2 = xDate.Text
    temptable!str3 = Format(XCODE.Text)
    temptable!str4 = xCodeDesca.Caption
'    temptable!str5 = xStore.Text
    temptable!Str11 = TurnValue(grid1.TextMatrix(I, 1))
    
    temptable!str12 = TurnValue(grid1.TextMatrix(I, 2))
    temptable!val1 = Val(grid1.TextMatrix(I, 3))
    temptable!val10 = I
    
    
    temptable!val9 = myPublic
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\Print_Imp.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub CMDTRANS_Click()
If Not myValidTrans Then Exit Sub
cString = InputBox(" «—ÌŒ «· —ÕÌ· ", " —ÕÌ·  ﬂ·›… «” Ì—«œÌ…", xDate.Text)
If Not IsDate(cString) Then
    MsgBox IIf(Trim(cString) = "", " „  Ã«Â· «· —ÕÌ·", "«· «—ÌŒ €Ì— ’ÕÌÕ")
    Exit Sub
End If
xDateTransPur.Text = cString
If MsgBox("Â·  Êœ «·Õ›Ÿ ﬁ»· «· —ÕÌ·", vbYesNo + vbDefaultButton1) = vbYes Then
    If mysave Then MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ Ê”Ì „ «· —ÕÌ· «·¬‰"
End If
If myTrans Then
    MsgBox " „ «· —ÕÌ· »‰Ã«Õ"
    CmdUndo_Click
End If
End Sub
Private Sub Command2_Click()
Dim DATABLE As New ADODB.Recordset
DATABLE.Open "FILE7_60", con, adOpenKeyset, adLockOptimistic, adCmdTable
With DATABLE
    .MoveFirst
    Do While Not .EOF
        !store = RetZero(!store, 2)
        .Update
        .MoveNext
    Loop
End With
End Sub
Private Sub Command3_Click()
On Error GoTo myerror
con.BeginTrans
con.Execute "UPDATE FILE7_60 SET FILE7_60.STORE = " & addstring(xStore.BoundText) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute "INSERT INTO FILE7_60 ( SERIAL, DOC_NO, ITEM, Quant, Price, Discount, Totalfrgn, [NUMBER], IMP_NO, inv_no_c, type, isClosed, row, cost, Total, RATE1, PRICE1, RATE2, PRICE2, Store )" & _
            " SELECT FILE7_60.SERIAL, " & addstring(xDoc_No.Text) & ", FILE7_60.ITEM, FILE7_60.Quant, FILE7_60.Price, FILE7_60.Discount, FILE7_60.Totalfrgn, FILE7_60.NUMBER, FILE7_60.IMP_NO, FILE7_60.inv_no_c, FILE7_60.type, FILE7_60.isClosed, FILE7_60.row, FILE7_60.cost, FILE7_60.Total, FILE7_60.RATE1, FILE7_60.PRICE1, FILE7_60.RATE2, FILE7_60.PRICE2, FILE7_60H.store " & _
            " FROM FILE7_60 INNER JOIN FILE7_60H ON FILE7_60.DOC_NO = FILE7_60H.DOC_NO " & _
            " WHERE FILE7_60H.DATE = " & DateSq(xDate.Text) & " AND  FILE7_60H.DOC_NO <> " & MyParn(xDoc_No.Text)
con.Execute "DELETE FILE7_60.* FROM FILE7_60 INNER JOIN FILE7_60H ON FILE7_60.DOC_NO = FILE7_60H.DOC_NO  WHERE FILE7_60H.DATE = " & DateSq(xDate.Text) & " AND  FILE7_60H.DOC_NO <> " & MyParn(xDoc_No.Text)
con.Execute "DELETE FILE7_60H.* FROM FILE7_60H WHERE FILE7_60H.DATE = " & DateSq(xDate.Text) & " AND  FILE7_60H.DOC_NO <> " & MyParn(xDoc_No.Text)
Dim loctable As New ADODB.Recordset
loctable.Open "SELECT * FROM FILE7_60 WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text) & " ORDER BY STORE,ROW", con, adOpenStatic, adLockOptimistic, adCmdText
I = 0
Do Until loctable.EOF
    I = I + 1
    loctable!Row = I
    loctable!Serial = I
    loctable.Update
    loctable.MoveNext
Loop
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
cFile = "File7_60"
cFileHeader = "File7_60H"
cFileClient = "File4_10"
With grid1
    .Cols = 13
    .RowHeight(0) = 700
    .Editable = flexEDKbd
    .WordWrap = True
End With
cList = StrList("select * from file0_40 order by desca")
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT FILE7_60H.*,FILE7_20H.DOC_NO AS DOCPUR  FROM FILE7_60H LEFT JOIN FILE7_20H  ON FILE7_60H.DOC_NO = FILE7_20H.DOCIMP ORDER BY FILE7_60H.DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE0_60"
Set xcurrency.RowSource = DATA2
xcurrency.ListField = "Desca"
xcurrency.BoundColumn = "Code"

Set grid1.DataSource = data3
data3.ConnectionString = strCon
CmdNewInv_Click

Command1.Visible = bopt1
cmdCalcCost.Visible = bopt1
Frame7.Visible = bopt1

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 1 Then
    GrdDesc Row
End If
Calctotals
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
grid1.AddItem grid1.Rows - 1
MakeSerial
End Sub

Private Sub Grid1_BeforeSort(ByVal Col As Long, Order As Integer)
grid1.RemoveItem grid1.Rows - 1
End Sub

Private Sub Grid1_DblClick()
With grid1
    If .Row > 0 Then
        If .Col = 1 Or .Col = 2 Then
            If MsgBox(" —ÕÌ· ·ÿ»«⁄…  ﬂ  ··’‰›", vbYesNo) = vbYes Then
                If GetDesca("select item  from ADDPRINT where item = " & MyParn(.TextMatrix(.Row, 1))) <> "" Then
                    If MsgBox("«·’‰› „—Õ· „‰ ﬁ»· - ≈⁄«œ…  —ÕÌ· «·’‰› ", vbOKCancel) = vbOK Then
                        con.Execute " delete  from ADDPRINT where item = " & MyParn(.TextMatrix(.Row, 1)) & " and isnull(addprint.desca)"
                    Else
                        Exit Sub
                    End If
                End If
                con.Execute "Insert Into ADDPRINT(Item,Quant,isPrint) " & _
                    " Values(" & _
                    addstring(.TextMatrix(.Row, 1)) & "," & _
                    addvalue(.TextMatrix(.Row, 3)) & "," & _
                    "1" & _
                    ")"
            
            End If
        End If
    End If
End With
End Sub
Private Sub Grid1_EnterCell()
If grid1.Row = 0 Then Exit Sub
If (grid1.Col = 0 Or grid1.Col = 2 Or grid1.Col = 6 Or grid1.Col = 7 Or grid1.Col = 8 Or grid1.Col = 9) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbd
    If grid1.Col = 3 Then
        If bopt1 Then grid1.Editable = flexEDNone
    End If
    If grid1.Col = 4 Then
        If Val(grid1.TextMatrix(grid1.Row, 4)) = 0 Then
            grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row - 1, 4)
        End If
        If Val(grid1.TextMatrix(grid1.Row, 4)) = 0 Then
            grid1.TextMatrix(grid1.Row, 4) = grid1.TextMatrix(grid1.Row - 1, 4)
        End If
    
    End If
End If
If (grid1.Col = 1) Then
    SetKbLayout Lang_EN
End If

End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then
    grid1.AddItem "", grid1.Row
    MakeSerial grid1.Row - 1
End If
If KeyCode = 112 And grid1.Col = 1 Then
    ItemsLookupAll Me, Search3
End If
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With grid1
If KeyAscii = 13 Then
    If Col = 3 Then
        grid1.Row = Row + 1
        grid1.Col = IIf(Row = grid1.Rows - 2, 1, 3)
    ElseIf Col = 1 Then
        grid1.Col = 3
     End If
End If

End With
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
End If
End Sub


Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
XCODE.Text = ""
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "Select code ,DescA From FILE4_10"
    Generalarray(2) = "Order by code"
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·»Ì«‰"
    listarray(0, 1) = "(%%DESCA%%)"
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·»Ì«‰"
    GrdArray(1, 1) = 6000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub

Private Sub xCode_LostFocus()
xCodeDesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(XCODE.Text)) & ""
End Sub

Private Sub xDateTrans_Change()
CMDTRANS.Enabled = IsDate(xDateTrans.Text)
End Sub

Private Sub xDiscount_LostFocus()
Calctotals
End Sub
Private Function MYVALID() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


'If xStore.BoundText = "" Then
'    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
'    Exit Function
'End If

If XCODE.Text = "" Then
    MsgBox "·„ Ì „ «œŒ«· ≈”„ «·⁄„Ì·"
    Exit Function
End If

If (Not IsDate(xDateTrans.Text)) And xtrans.Value <> 0 Then
    MsgBox " «—ÌŒ «· —ÕÌ· €Ì— ”·Ì„ «Ê €Ì— „ÊÃÊœ"
    Exit Function
End If
With grid1
For I = 1 To grid1.Rows - 2
    If Not validRow(I) Then
        MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
        Exit Function
    End If
Next
If xtrans.Value <> 0 Then
    cString = Trim(GetDesca("Select Doc_no from file7_20H where docimp = " & MyParn(xDoc_No.Text)) & "")
    If cString <> "" Then
        MsgBox "·‰ Ì „ «·Õ›Ÿ «·—Ã«¡ «€·«ﬁ «Œ Ì«— «· —ÕÌ· ÕÌÀ «‰Â  „  —ÕÌ· «·„” ‰œ „‰ ﬁ»·" & vbCrLf & _
               " „” ‰œ „‘ —Ì«  —ﬁ„ : " & cString
        Exit Function
    End If
End If
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "YYYY-MM-DD")
'xStore.BoundText = CardTable!Store

XCODE.Text = CardTable!Code
xtrans.Value = IIf(CardTable!Trans, 1, 0)
xCredit.Text = CardTable!CREDIT & ""
xVessel.Text = CardTable!Vessel & ""
xBankName.Text = CardTable!BankName & ""
xPolicy.Text = CardTable!Policy & ""
xcurrency.BoundText = CardTable!Currency & ""
xDateTransPur.Text = Format(CardTable!DateTransPur, "YYYY-MM-DD")
xcurRate.Text = CardTable!curRate & ""
XFACTNAME.Text = CardTable!FACTNAME & ""
xCodeDesca.Caption = GetDesca("select desca from " & cFileClient & " where code = " & MyParn(XCODE.Text)) & ""
xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = Format(CardTable!Discount & "", "Fixed")
xcharge.Caption = GetDesca("Select sum([value]) from fiLE7_60CH WHERE DOC_NO = " & MyParn(xDoc_No.Text))
CMDTRANS.Enabled = IsNull(CardTable!docpur)
cmduntrans.Enabled = Not IsNull(CardTable!docpur)
cString = "SELECT FILE7_60.ROW,FILE1_10.ITEM,FILE1_10.DESCA,FILE7_60.Quant,FILE7_60.Price,FILE7_60.DISCOUNT,0 AS TOTALFRGN,FILE7_60.TOTAL,FILE7_60.COST,FILE7_60.RATE1,FILE7_60.PRICE1,FILE7_60.RATE2,FILE7_60.PRICE2,FILE7_60.STORE,FILE7_60.ID" & _
          " FROM FILE7_60 INNER JOIN FILE1_10 ON FILE7_60.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & _
          " ORDER BY FILE7_60.ROW"
data3.RecordSource = cString
data3.Refresh
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
Handlecontrols LoadMode
Calctotals
Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag("FILE7_60H", "doc_no")), 15)
xDate.Text = ""
xCredit.Text = ""
xVessel.Text = ""
xtrans.Value = 0
xBankName.Text = ""
xPolicy.Text = ""
xcurrency.BoundText = ""
xDateTransPur.Text = ""
xcurRate.Text = ""
XFACTNAME.Text = ""
xStore.BoundText = ""

xCodeDesca.Caption = ""
XCODE.Text = ""
xDiscount.Text = ""
xTotal.Caption = ""
xTotalItem.Caption = ""
xusername.Text = ""
CMDTRANS.Enabled = False
cmduntrans.Enabled = False
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(1, 0) = 1
Fixgrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bEdit
'If Not (CardTable.EOF And CardTable.BOF) Then cmdTrans.Enabled = IsNull(CardTable!docPur) And nMode = LoadMode Else cmdTrans.Enabled = False
'If Not (CardTable.EOF And CardTable.BOF) Then cmduntrans.Enabled = (Not IsNull(CardTable!docPur)) And nMode = LoadMode And bEdit Else cmdTrans.Enabled = False
'xTrans.Enabled = Not cmduntrans.Enabled

cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = (nMode = LoadMode) And bEdit
'cmdSave.Enabled = (bEdit) And (CanEdit) Or nMode = DefineMode
'CmdDelInv.Enabled = (nMode = LoadMode And CanEdit) And bDel
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text, 15)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ", vbOKCancel + vbDefaultButton2) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then con.Execute "DELETE FROM FILE7_60 WHERE ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        grid1.RemoveItem grid1.Row
        Calctotals
        CalcCost
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 1
        If KeyCode = 27 Then Exit Sub
End Select
End Sub
Private Sub GrdDesc(Row)
If grid1.Col <> 1 Then Exit Sub
For nCol = 2 To grid1.Cols - 2
    grid1.TextMatrix(Row, nCol) = ""
Next
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
grid1.TextMatrix(Row, 2) = GetDesca("select desca from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 1)))
'If grid1.Row > 1 Then If SameGroup(grid1.TextMatrix(Row, 1), grid1.TextMatrix(Row - 1, 1)) Then grid1.TextMatrix(Row, 4) = grid1.TextMatrix(Row - 1, 4)
End Sub
Private Function Calctotals()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalFrgn As Double, nTotalNoCharge
With grid1
For I = 1 To grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    grid1.TextMatrix(I, 6) = Round(Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount, 2)
    nTotalitem = nTotalitem + Val(grid1.TextMatrix(I, 6))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 3))
Next
nTotalFrgn = nTotalitem - Val(xDiscount.Text)
nTotalNoCharge = nTotalFrgn * Val(xcurRate.Text)
xTotalItem.Caption = Format(nTotalitem, "Fixed")
xTotalFrgn.Caption = Format(nTotalFrgn, "Fixed")
xTotalNoCharge.Caption = Format(nTotalNoCharge, "Fixed")
xTotal.Caption = (nTotalFrgn * Val(xcurRate.Text)) + Val(xcharge.Caption)
xtotalQuant.Caption = Format(nTotalQuant, "#0.0000")
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(7, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DISTINCT fILE7_60H.DOC_NO,FILE7_60H.DATE,CONVERT(VARCHAR(10),FILE7_60H.DATE,111),FILE4_10.Desca,FILE7_60H.Credit,FILE7_60H.FactName, CASE WHEN (FILE7_20H.DOCIMP is Null) THEN '€Ì— „—Õ·' ELSE '„—Õ·' END" & _
                  " FROM ((FILE7_60H left JOIN FILE4_10 ON FILE7_60H.CODE = FILE4_10.CODE) LEFT JOIN FILE7_20H ON FILE7_60H.DOC_NO = FILE7_20H.DOCIMP)"

Generalarray(2) = "Order by FILE7_60H.Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ-≈”„ «·—”«·…"
listarray(0, 1) = "(@@FILE7_60H.Doc_No@@15 Or  %%FILE4_10.DESCA%% OR %%FILE7_60H.FactName%% OR " & _
                  "##FILE7_60H.Date##)"

listarray(1, 0) = "—ﬁ„ «·«⁄ „«œ"
listarray(1, 1) = "(FILE7_60H.Credit Like '%cFilter%')"

listarray(2, 0) = "«·„Œ“‰"
listarray(2, 1) = "(FILE7_60H.STORE = 'cFilter')"
listarray(2, 2) = "FILE0_40"
listarray(2, 3) = "CODE"
listarray(2, 4) = "DESCA"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ «·«⁄ „«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«”„ «·„—ﬂ»"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "«· —ÕÌ·"
GrdArray(6, 1) = 1100


searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow And Trim(grid1.TextMatrix(I, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 1)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(grid1.Row, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function
Private Sub xfilter_GotFocus(Index As Integer)
SetKbLayout Lang_EN
xfilter(Index).SelStart = 0
xfilter(Index).SelLength = Len(xfilter(Index).Text)
End Sub

Private Sub xfilter_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xfilter(Index), Index
End If
End Sub
Private Sub xRate_LostFocus()
Calctotals
End Sub
Private Function validRow(nRow) As Boolean
If nRow > 0 Then
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
    If Trim(grid1.TextMatrix(nRow, 2)) = "" Then Exit Function
   ' If Val(grid1.TextMatrix(nRow, 5)) = 0 Then Exit Function
End If
validRow = True
End Function
Sub additemProc()
grid1.RemoveItem grid1.Rows - 1
With additemfrm.grid1
    For I = 1 To .Rows - 1
        If Val(.TextMatrix(I, 4)) <> 0 Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(I, 0)
            grid1.TextMatrix(grid1.Rows - 1, 2) = retitem(.TextMatrix(I, 0), "desca")
            grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(I, 4)
            grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(I, 5)
            grid1.TextMatrix(grid1.Rows - 1, 7) = retitem(.TextMatrix(I, 0), "width1") & ""
            grid1.TextMatrix(grid1.Rows - 1, 8) = retitem(.TextMatrix(I, 0), "width2") & ""
            grid1.TextMatrix(grid1.Rows - 1, 9) = retitem(.TextMatrix(I, 0), "length") & ""
        End If
    Next
    grid1.AddItem ""
   Calctotals
End With
End Sub
Private Sub CalcCost()
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalFrgn As Double, nTotalNoCharge
With grid1
For I = 1 To grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    grid1.TextMatrix(I, 6) = Round(Val(.TextMatrix(I, 3)) * Val(.TextMatrix(I, 4)) * nDiscount, 2)
    nTotalitem = nTotalitem + Val(grid1.TextMatrix(I, 6))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 3))
Next
nTotalFrgn = nTotalitem - Val(xDiscount.Text)
nTotalNoCharge = nTotalFrgn * Val(xcurRate.Text)
nTotal = nTotalNoCharge + Val(xcharge.Caption)

If Val(nTotalFrgn) = 0 Then Exit Sub
nRate = nTotal / nTotalFrgn
For I = 1 To grid1.Rows - 2
    nDiscount = 1 - (Val(.TextMatrix(I, 5)) / 100)
    grid1.TextMatrix(I, 8) = Round(Val(grid1.TextMatrix(I, 4)) * nDiscount * nRate, 2)
    grid1.TextMatrix(I, 7) = Round(Val(grid1.TextMatrix(I, 8)) * Val(grid1.TextMatrix(I, 3)), 2)
Next
End With
End Sub
Private Sub doprint()
Dim aHeader(3)
If Val(xTotalFrgn.Caption) = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ’ÕÌÕ… ·ÿ»«⁄ Â«"
    Exit Sub
End If
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset

contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.[GROUP],FILE7_60.ITEM,File1_10.Desca,FILE7_60.CODE, FILE7_60.Quant, FILE7_60.COST,FILE7_60.PRICE,FILE7_60.PRICE1,FILE7_60.PRICE2  " & _
          " FROM FILE1_10 INNER JOIN FILE7_60 ON FILE1_10.ITEM = FILE7_60.ITEM" & _
          " WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text) & _
          " ORDER BY FILE7_60.ROW "
         
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

aHeader(0) = "[" & "›« Ê—… —ﬁ„ : " & xDoc_No.Text & "]"
aHeader(1) = "[" & "» «—ÌŒ : " & xDate.Text & "]"
aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
If xVessel.Text <> "" Then aHeader(3) = "[" & " «·”›Ì‰… : " & xVessel.Text & "]"
With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!val10 = Val(xTotalFrgn.Caption)
        temptable!val11 = Val(xcurRate.Text)
        temptable!val12 = Val(xcharge.Caption)
        temptable!val13 = Val(xTotal.Caption)
        temptable!VAL14 = (Val(xTotal.Caption)) / Val(xTotalFrgn.Caption)
        temptable!str1 = !Item
        temptable!str2 = !Desca
        temptable!val1 = !Quant
        temptable!val2 = !price
        temptable!val3 = Val(!Quant & "") * Val(!price & "")
        temptable!val4 = !cost
        temptable!val5 = Val(!Quant & "") * Val(!cost & "")
        temptable!Val6 = !Price1
        temptable!Val7 = Val(!Quant & "") * Val(!Price1 & "")
        temptable!Val8 = !PRICE2
        temptable!val9 = Val(!Quant & "") * Val(!PRICE2 & "")
        temptable!str21 = retHeader(aHeader, 0, 2)
        temptable!str22 = retHeader(aHeader, 2, 2)
        temptable.Update
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    Report1.ReportFileName = App.Path & "\Reports\impcost_1.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    Report1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbd
    .FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·Œ’„|" & "«·≈Ã„«·Ì »«·⁄„·…|" & "≈Ã„«·Ì «· ﬂ·›… »«·Ã‰ÌÂ|" & " ﬂ·›… «·ÊÕœ… »«·Ã‰ÌÂ|" & "‰”»… «·Ã„·…|" & "”⁄— «·Ã„·…|" & "‰”»… „” Â·ﬂ|" & "”⁄— „” Â·ﬂ|" & "«·„Œ“‰|"
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 4000
    .ColWidth(3) = 800
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 800
    .ColWidth(13) = 2000
    .ColComboList(13) = cList
    .ColHidden(grid1.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    .ColHidden(4) = Not bopt1
    .ColHidden(5) = Not bopt1
    .ColHidden(6) = Not bopt1
    .ColHidden(7) = Not bopt1
    .ColHidden(8) = Not bopt1
    .ColHidden(9) = Not bopt1
    .ColHidden(10) = Not bopt1
    .ColHidden(11) = Not bopt1
    .ColHidden(12) = Not bopt1
    .ExplorerBar = flexExSortShow
End With
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = nBeginRow To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Function myunTrans() As Boolean
On Error GoTo myerror
con.BeginTrans
con.Execute "DELETE FILE7_20 FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO" & _
        " WHERE FILE7_20H.DOCIMP = " & MyParn(xDoc_No.Text)

con.Execute "DELETE  FROM FILE7_20H " & _
        " WHERE FILE7_20H.DOCIMP = " & MyParn(xDoc_No.Text)
con.CommitTrans
myunTrans = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function myTrans() As Boolean
Dim nTry As Integer, StoresTable As New ADODB.Recordset
StoresTable.Open "Select Store From file7_60 where Doc_no = " & MyParn(xDoc_No.Text) & " Group by Store order by Store", con, adOpenStatic, adLockReadOnly, adCmdText
On Error GoTo myerror
con.BeginTrans
Do Until StoresTable.EOF
    cDoc_no = RetZero(Newflag("FILE7_20H", "doc_no", con), 6)
    con.Execute "INSERT INTO FILE7_20h (DOC_NO, code, store, [date], discount,[type],DocImp) " & _
                " SELECT " & MyParn(cDoc_no) & ", FILE7_60H.code, " & MyParn(StoresTable!store) & ", " & addDate(xDateTransPur.Text) & ", FILE7_60H.discount,1,FILE7_60H.DOC_NO" & _
                " FROM FILE7_60H " & _
                " WHERE FILE7_60H.DOC_NO = " & MyParn(xDoc_No.Text)
    
    con.Execute "INSERT INTO FILE7_20 ( DOC_NO, ITEM, Quant, PRICE,[ROW]) " & _
          " SELECT " & MyParn(cDoc_no) & ", FILE7_60.ITEM, FILE7_60.Quant, FILE7_60.cost ,[ROW] " & " FROM FILE7_60 " & _
          " WHERE FILE7_60.DOC_NO = " & MyParn(xDoc_No.Text) & " And store = " & MyParn(StoresTable!store)
    StoresTable.MoveNext
Loop
con.CommitTrans
myTrans = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function

Private Function CanEdit() As Boolean
If Not (CardTable.EOF And CardTable.BOF) Then If Not IsNull(CardTable!docpur) Then Exit Function
CanEdit = True
End Function
Private Function myValidTrans() As Boolean
Dim cString
cString = GetDesca("Select Doc_no from file7_20H where docimp = " & MyParn(xDoc_No.Text)) & ""
If Trim(cString) <> "" Then
    MsgBox "Â‰«ﬂ „” ‰œ „‘ —Ì«   „  —ÕÌ·… »—ﬁ„ " & cString
    Exit Function
End If
myValidTrans = True
End Function
Private Sub doprint2()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(0)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA AS ITEMDESCA,FILE1_10.UNIT, " & _
          "  FILE7_60.DOC_NO,FILE7_60H.CODE as SupCode,FILE7_60H.DATE,FILE7_60H.VESSEL, " & _
          " Sum(val(FILE7_60.Quant & '')) AS SumofQuant, " & _
          " FILE1_10.[GROUP] AS GroupCode, FILE1_50.DESCA AS GroupDesca,  " & _
          " FILE1_50.[GROUP] AS MainGroupCode, FILE1_51.DESCA as  MainGroupDesca" & _
          " FROM (((FILE7_60 INNER JOIN FILE7_60H ON FILE7_60.DOC_NO = FILE7_60H.DOC_NO )INNER JOIN FILE1_10  ON FILE7_60.ITEM = FILE1_10.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_51 ON " & _
          " FILE1_50.[GROUP] = FILE1_51.CODE"

cString = cString & turnFound(cString) & "FILE7_60H.doc_no = " & MyParn(xDoc_No.Text)
cString = cString & " Group by  FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA,FILE1_10.UNIT," & _
          " FILE7_60.DOC_NO,FILE7_60H.CODE,FILE7_60H.DATE,FILE7_60H.VESSEL, " & _
          " FILE1_10.[GROUP], FILE1_50.DESCA,FILE1_50.[GROUP], FILE1_51.DESCA  "

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!str6 = !MainGroupDesca
        temptable!str5 = !MAINGROUPCODE
        temptable!str1 = !Item
        temptable!str2 = itemWidth(sourcetable!Item)
        temptable!str3 = !GroupCode
        temptable!str4 = !GroupDesca
        temptable!str8 = GetDesca("Select Desca from file1_13 where code = " & MyParn(!UNIT))
        temptable!str9 = !doc_no
        temptable!str10 = GetDesca("Select Desca from file4_10 where code = " & MyParn(!supCode))
        temptable!Str11 = !Vessel
        temptable!Date1 = !Date
        temptable!val1 = !sumOfQuant
        temptable!Val20 = !width1
        temptable!STR20 = !width1
        temptable!str17 = TurnValue(retHeader(aHeader, 0, 4))
        temptable.Update
        .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    Report1.ReportFileName = App.Path & "\Reports\invdtl.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    Report1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Function itemWidth(pItem) As String
itemWidth = retitem(pItem, "width1") & ""
If Not IsNull(retitem(pItem, "width2")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "width2")
If Not IsNull(retitem(pItem, "length")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "length")
End Function
Function mysave(Optional bIgMsg As Boolean = False)
If Not MYVALID Then Exit Function
CalcCost
Calctotals
If Not myreplace Then Exit Function
CardTable.Requery
If Not bIgMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
myload
mysave = True
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Cancel = True
    End If
    If GetDesca("select item from file1_10 where item = " & MyParn(grid1.EditText)) = "" Then
        If grid1.EditText <> "" Then MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
        Cancel = True
        Exit Sub
    End If
End If
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub myreplaceGrd()
Dim aGrid(13, 1)
prog1.Value = 0
prog1.Visible = True
With grid1
    For I = 1 To .Rows - 1
        prog1.Value = Round(I / (grid1.Rows), 2) * 100
        aGrid(0, 0) = "Doc_no"
        aGrid(0, 1) = addstring(xDoc_No.Text)
        
        aGrid(1, 0) = "item"
        aGrid(1, 1) = addstring(grid1.TextMatrix(I, 1))
        
        aGrid(2, 0) = "Quant"
        aGrid(2, 1) = Val(grid1.TextMatrix(I, 3))
        
        aGrid(3, 0) = "Price"
        aGrid(3, 1) = Val(grid1.TextMatrix(I, 4))
        
        aGrid(4, 0) = "Discount"
        aGrid(4, 1) = Val(grid1.TextMatrix(I, 5))
        
        aGrid(5, 0) = "TotalFrgn"
        aGrid(5, 1) = Val(grid1.TextMatrix(I, 6))
        
        aGrid(6, 0) = "Total"
        aGrid(6, 1) = Val(grid1.TextMatrix(I, 7))
        
        aGrid(7, 0) = "Cost"
        aGrid(7, 1) = Val(grid1.TextMatrix(I, 8))
        
        aGrid(8, 0) = "Rate1"
        aGrid(8, 1) = Val(grid1.TextMatrix(I, 9))
        
        aGrid(9, 0) = "Price1"
        aGrid(9, 1) = Val(grid1.TextMatrix(I, 10))
        
        aGrid(10, 0) = "RATE2"
        aGrid(10, 1) = Val(grid1.TextMatrix(I, 11))
        
        aGrid(11, 0) = "PRICE2"
        aGrid(11, 1) = Val(grid1.TextMatrix(I, 12))
        
        aGrid(12, 0) = "STORE"
        aGrid(12, 1) = addstring(grid1.TextMatrix(I, 13))
        
        aGrid(13, 0) = "ROW"
        aGrid(13, 1) = I
        
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aGrid, "FILE7_60")
        Else
            con.Execute CreateUpdate(aGrid, "FILE7_60", " where ID = " & grid1.TextMatrix(I, grid1.Cols - 1), Array(-1))
        End If
    Next
    prog1.Visible = False
End With
End Sub
Private Sub xusername_GotFocus()
xusername.SelStart = 0
xusername.SelLength = Len(xusername.Text)
End Sub
Private Sub xDateTransPur_GotFocus()
xDateTransPur.SelStart = 0
xDateTransPur.SelLength = Len(xDateTransPur.Text)
End Sub
Private Sub xDateTrans_GotFocus()
xDateTrans.SelStart = 0
xDateTrans.SelLength = Len(xDateTrans.Text)
End Sub
Private Sub xcurRate_GotFocus()
xcurRate.SelStart = 0
xcurRate.SelLength = Len(xcurRate.Text)
End Sub
Private Sub xFactName_GotFocus()
XFACTNAME.SelStart = 0
XFACTNAME.SelLength = Len(XFACTNAME.Text)
End Sub
Private Sub xPolicy_GotFocus()
xPolicy.SelStart = 0
xPolicy.SelLength = Len(xPolicy.Text)
End Sub
Private Sub xBankName_GotFocus()
xBankName.SelStart = 0
xBankName.SelLength = Len(xBankName.Text)
End Sub
Private Sub xCredit_GotFocus()
xCredit.SelStart = 0
xCredit.SelLength = Len(xCredit.Text)
End Sub
Private Sub xVessel_GotFocus()
xVessel.SelStart = 0
xVessel.SelLength = Len(xVessel.Text)
End Sub
Private Sub xCode_GotFocus()
XCODE.SelStart = 0
XCODE.SelLength = Len(XCODE.Text)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDiscount_GotFocus()
xDiscount.SelStart = 0
xDiscount.SelLength = Len(xDiscount.Text)
End Sub

