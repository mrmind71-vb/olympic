VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form Member_hfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄÷ÊÌ… ‘—›ÌÂ"
   ClientHeight    =   5115
   ClientLeft      =   690
   ClientTop       =   1395
   ClientWidth     =   10350
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "member_h.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   10350
   Begin VB.Frame Frame1 
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
      Height          =   3435
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   7890
      Begin VB.CheckBox xCurrent_mem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄÷Ê Õ«·Ì"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2925
         Width           =   1230
      End
      Begin VB.TextBox xNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4725
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   270
         Width           =   1275
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   1125
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1710
         Width           =   5910
      End
      Begin VB.CheckBox xFamily 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ê«·⁄«∆·…"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1185
      End
      Begin VB.TextBox xDate_end 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4410
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1350
         Width           =   1590
      End
      Begin VB.TextBox xTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3330
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   2670
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   5910
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3330
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label xcode_zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   270
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1755
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ…"
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1395
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«··ﬁ»"
         Height          =   270
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   285
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·«”„"
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   645
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   7890
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3960
         TabIndex        =   6
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":0342
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":2D37
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":55D0
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":78F3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":A08F
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2655
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":C523
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":E764
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   5265
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":10A51
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":12A59
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   6570
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":14A10
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":16DDB
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1800
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   -2295
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   7110
      Top             =   7515
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
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "member_h.frx":18E84
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   12330
      TabIndex        =   11
      Tag             =   "-1"
      Top             =   495
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind71"
      LicenseRegCode  =   "íß“Ωª∫≠Ω≥´±“™ºØ´¥æÆØUBOR-FEOEONZI-EPCP6gI"
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   4740
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   45
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   476
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   3465
         TabIndex        =   29
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   6975
         TabIndex        =   30
         Top             =   45
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCommand cmdDelIInstall 
      Height          =   510
      Left            =   8640
      TabIndex        =   39
      Top             =   4230
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
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
      Caption         =   "÷»ÿ «·«—ﬁ«„"
      ButtonStyle     =   3
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   2220
      Begin Threed.SSCommand cmdScan 
         Height          =   555
         Left            =   1125
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
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
         Picture         =   "member_h.frx":192C6
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":1B7CE
      End
      Begin Threed.SSCommand cmdDelPhoto 
         Height          =   555
         Left            =   90
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   180
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
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
         Picture         =   "member_h.frx":1E369
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":2092E
      End
      Begin Threed.SSCommand cmdAddCatalog 
         Height          =   555
         Left            =   90
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "member_h.frx":233FF
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_h.frx":25942
      End
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         Height          =   2595
         Left            =   90
         Stretch         =   -1  'True
         Top             =   765
         Width           =   2025
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4095
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "member_h.frx":27C40
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_h.frx":29DE7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   180
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         Picture         =   "member_h.frx":2BE2E
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_h.frx":2DF19
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   741
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
         Picture         =   "member_h.frx":2FF13
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_h.frx":32024
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "member_h.frx":3401E
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "member_h.frx":36242
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4095
      Width           =   4515
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬂ· «·«⁄÷«¡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   225
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "”‰Ê«  ”«»ﬁ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Õ«·ÌÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Member_hfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection, bCheck As Boolean
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nPhoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String
Dim formMode As Byte
Dim oSearch As New Search
Dim CardTable As ADODB.Recordset
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
cmdDel.Enabled = (nMode = LoadMode And bEditRecord)
cmdSave.Enabled = bEditRecord
cmdInform.Enabled = (nMode = LoadMode)
cmdScan.Enabled = nMode = LoadMode And bEditRecord

aRecords = retRecords(xcode_zero.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    panel1(0).Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)
Else
    panel1(0).Caption = ArbString("«÷«›… ”Ã· " & (nRecords + 1))
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sCode = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sCode = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sCode = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sCode = ""
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = Newflag("FILE3_10", "code", con)
xcode_zero.Caption = ""
xNo.Text = ""
xDesca.Text = ""
xTitle.Text = ""
xDate_End.Text = ""
xNotes.Text = ""
xFamily.Value = 0
xCurrent_mem.Value = 1
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
End If
myUndo
End Sub
Private Sub myload(Optional bNoGrid As Boolean = False)
xCode.Text = CardTable!CODE & ""
xcode_zero.Caption = CardTable!CODE_ZERO & ""
xNo.Text = CardTable!NO & ""
xDesca.Text = CardTable!desca & ""
xTitle.Text = CardTable!Title & ""
xDate_End.Text = myFormat_p(CardTable!date_end)
xNotes.Text = CardTable!notes & ""
bCheck = True
xCurrent_mem.Value = IIf(CardTable!current_mem, 1, 0)
xFamily.Value = IIf(CardTable!FAMILY, 1, 0)
bCheck = False
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
If validPhoto(RetPhotoh(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhotoh(xCode.Text))
End Sub
Private Function MyReplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[NO]", addvalue(xNo.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "[NOTES]", addstring(xNotes.Text))
aInsert = AddFlag(aInsert, "[TITLE]", addstring(xTitle.Text))
aInsert = AddFlag(aInsert, "[DATE_END]", addDate(xDate_End.Text))
aInsert = AddFlag(aInsert, "[FAMILY]", xFamily.Value)
aInsert = AddFlag(aInsert, "[current_mem]", xCurrent_mem.Value)
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.Text))
    con.Execute addInsert(aInsert, "FILE3_10")
Else
    con.Execute addUpdate(aInsert, "FILE3_10", "FILE3_10.CODE = " & xCode.Text)
End If
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub cmdAdd_Click()
mydefine
xDesca.SetFocus
End Sub

Private Sub cmdAddCatalog_Click()
If Trim(xCode.Text) = "" Then Exit Sub
Set fs = CreateObject("Scripting.FileSystemObject")
Dim cFile As String, cNewFile As String
On Error GoTo myerror
Common1.FileName = ""
Common1.InitDir = App.Path
Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
Common1.ShowOpen
If Common1.FileTitle <> "" Then
    cFile = Common1.FileName
    If cFile <> "" Then
        cNewFile = RetPhotoh(xCode.Text)
        fs.CopyFile cFile, cNewFile
    End If
    xMemberPhoto.Picture = LoadPicture("")
    If validPhoto(RetPhotoh(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhotoh(xCode.Text))
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CmdDel_Click()
Dim nDelete As Long
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From FILE3_10 Where code = " & xCode.Text, nDelete
    Dim fs As New FileSystemObject
    If fs.FileExists(RetPhotoh(xCode.Text)) Then
        fs.DeleteFile RetPhotoh(xCode.Text)
    End If
    con.CommitTrans
    
    If nDelete = 0 Then
        MsgBox "·„ Ì „ «·‰Ÿ«„ „‰ Õ–› «·⁄÷Ê"
        Exit Sub
    End If
    DeletePhoto xCode.Text
    openCardTable xCode.Text, "<="
    If CardTable.EOF Then openCardTable , ">"
    If CardTable.EOF Then
        mydefine
    Else
        myload
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub

Private Sub cmdDelImage_Click()
End Sub

Private Sub cmdDelIInstall_Click()
If MsgBox(" ⁄œÌ· «·«—ﬁ«„ ?", vbOKCancel) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myerror
con.Execute "update file3_10 set file3_10.no = dbo.f_serial_no(FILE3_10.CODE_ZERO) WHERE (NOT FILE3_10.NO IS NULL)"
con.CommitTrans
Inform " „ «· ⁄œÌ· »‰Ã«Õ"
myUndo
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
MemberH_LookupAll Me, oSearch
End Sub
Private Sub cmdclass_Click()
Dim oFlagfrm As New flag_mainfrm, sBoundText As String
sBoundText = xClass.BoundText
oFlagfrm.sTable = "class_CODES"
oFlagfrm.sCaption = "«·‘⁄»"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
Set data1.Recordset = myRecordSet("select * from class_Codes", con)
xClass.BoundText = sBoundText
If Not xClass.MatchedWithList Then xClass.BoundText = ""
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
myUndo
End Sub
Private Sub CmdNext_Click()
openCardTable xcode_zero.Caption, ">"
If CardTable.EOF Then openCardTable xCode.Text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xcode_zero.Caption, "<"
If CardTable.EOF Then openCardTable xCode.Text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub cmdScan_Click()
nOption = 0
ScanImage
If validPhoto(RetPhotoh(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhotoh(xCode.Text))
End Sub
Private Sub CmdUndo_Click()
myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()

openCon con
bEdit = True

myUndo
End Sub
Private Sub cmdDelPhoto_Click()
On Error GoTo myerror
If MsgBox("Õ–› «·’Ê—… Â· «‰  „ √ﬂœ ?", vbOKCancel) = vbOK Then
    If fs.FileExists(RetPhotoh(xCode.Text)) Then fs.DeleteFile RetPhotoh(xCode.Text)
    imgx1.Images.Clear
    xMemberPhoto.Picture = LoadPicture("")
End If
Exit Sub
myerror:
Err.Clear
MsgBox Err.Description
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Option1_Click(Index As Integer)
myUndo
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.Text) Then
     If xCode.Tag = LoadMode Then
        mydefine
    Else
        xCode.Text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
        If CardTable!CODE = xCode.Text Then
            Exit Sub
        End If
    End If
    
    openCardTable xCode.Text
    If Not CardTable.EOF Then
        myload
    ElseIf xCode.Tag = LoadMode Then
        mydefine
    Else
        'xCode.Text = ""
    End If
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not ValidInt(xCode.Text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·ÿ«·» €Ì— „”Ã·", , systemName
    Exit Function
End If

'If Not xClass.MatchedWithList Then
'    MsgBox "«·‘⁄»… €Ì— „”Ã·…"
'    Exit Function
'End If

If xDesca.Text = "" Then
    If Not igMsg Then MsgBox "√”„ «·ÿ«·» €Ì— „”Ã·", , systemName
    Exit Function
End If
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE3_10.* FROM FILE3_10"
If pSign = "=" Then
    If pCode <> "" Then cWhere = "CODE  " & pSign & addstring(pCode)
Else
    If pCode <> "" Then cWhere = "CODE_ZERO " & pSign & addstring(pCode)
End If

cFilter = ""
If Option1(1).Value Then
    cFilter = "FILE3_10.current_mem = 1"
ElseIf Option1(2).Value Then
    cFilter = "FILE3_10.current_mem = 0"
End If
If sCode <> "" Then cFilter = "FILE3_10.CODE = " & addvalue(sCode)
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by FILE3_10.CODE_ZERO desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by FILE3_10.CODE_ZERO ASC"
End If

CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Text) Then
    openCardTable xCode.Text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
On Error Resume Next
xDesca.SetFocus
Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
If twain.Sources.Count > 1 Then twain.SelectSource
twain.Acquire False, Me.hwnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
If Not ValidInt(xCode.Text) Then Exit Sub
ReplaceFromImage Image, RetPhotoh(xCode.Text)
xMemberPhoto.Picture = LoadPicture(RetPhotoh(xCode.Text))
End Sub
Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
' This event is called after you call Acquire.
' It let's you know when it's safe to call CloseTwain.
twain.CloseTwain
' Steps menu
End Sub
Private Sub ReplaceFromImage(Image As ImgX_Image, cPhoto)
On Error GoTo myerror
imgx1.Images.Replace Image, , False
imgx1.Refresh
imgx1.Export.ToFile cPhoto, ixfsJPG
Exit Sub
myerror:
imgx1.Images.Clear
Err.Clear
End Sub
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If Trim(pCode) <> "" Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE_ZERO <= " & MyParn(pCode) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM FILE3_10 " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If Trim(pCode) <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub xCurrent_mem_Click()
If Not bCheck Then
    If xCurrent_mem.Value = 1 Then
        xNo.Text = mRound(GetField("select top 1 no from file3_10 where current_mem = 1 order by no desc", con) + 1)
    Else
        xNo.Text = ""
    End If
End If
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_end_GotFocus()
myGotFocus xDate_End
End Sub
Private Sub xDate_end_LostFocus()
myLostFocus xDate_End
myValidDate xDate_End
End Sub
Private Sub xTitle_GotFocus()
myGotFocus xTitle
End Sub
Private Sub xTitle_LostFocus()
myLostFocus xTitle
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub


