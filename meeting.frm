VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form meetingFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·Ã„⁄Ì… «·⁄„Ê„Ì…"
   ClientHeight    =   5265
   ClientLeft      =   690
   ClientTop       =   1395
   ClientWidth     =   8115
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   8115
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
      Height          =   3300
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   7890
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈€·«ﬁ «·Ã„⁄Ì…"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2160
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4410
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2160
         Width           =   1590
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2520
         Width           =   5910
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4410
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   1590
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   5910
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4635
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1365
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Left            =   3015
         TabIndex        =   27
         Top             =   1350
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«Œ «— «·„Ê”„"
         Alignment       =   4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYearPaid 
         Height          =   375
         Left            =   3015
         TabIndex        =   35
         Top             =   1755
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«Œ «— «·„Ê”„"
         Alignment       =   4
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ì”œœ ›Ï „Ê”„"
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
         Index           =   1
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1845
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„”œœ Õ Ï"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2205
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„”œœ „Ê”„"
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
         Index           =   4
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label xcode_zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
         Height          =   285
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2610
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·Ã„⁄Ì…"
         Height          =   330
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1035
         Width           =   1185
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·ﬂÊœ"
         Height          =   285
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ì«‰ «·Ã„⁄Ì…"
         Height          =   330
         Index           =   0
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   1320
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   7890
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3960
         TabIndex        =   4
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
         Picture         =   "meeting.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "meeting.frx":29F5
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   10
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
         Picture         =   "meeting.frx":528E
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1350
         TabIndex        =   11
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
         Picture         =   "meeting.frx":75B1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "meeting.frx":9D4D
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2655
         TabIndex        =   12
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
         Picture         =   "meeting.frx":C1E1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "meeting.frx":E422
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   5265
         TabIndex        =   13
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
         Picture         =   "meeting.frx":1070F
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "meeting.frx":12717
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   6570
         TabIndex        =   14
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
         Picture         =   "meeting.frx":146CE
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "meeting.frx":16A99
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4185
      Top             =   4185
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
      DragIcon        =   "meeting.frx":18B42
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   12330
      TabIndex        =   8
      Tag             =   "-1"
      Top             =   495
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   4890
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4005
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   17
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
         Picture         =   "meeting.frx":18F84
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "meeting.frx":1B12B
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   18
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
         Picture         =   "meeting.frx":1D172
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "meeting.frx":1F25D
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   19
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
         Picture         =   "meeting.frx":21257
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "meeting.frx":23368
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   20
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
         Picture         =   "meeting.frx":25362
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "meeting.frx":27586
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   4005
      Width           =   2220
      Begin Threed.SSCommand cmdAddMember 
         Height          =   420
         Left            =   45
         TabIndex        =   33
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«÷«›… √⁄÷«¡ «·Ã„⁄Ì…"
         ButtonStyle     =   3
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   34
      Top             =   4740
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   0
   End
End
Attribute VB_Name = "meetingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection, bCheck As Boolean, oSearchYear As New Search_empty
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
bEditRecord = bedit
bEditRecord = xClosed.Value = 0

cmdAddMember.Enabled = bEditRecord And nMode = LoadMode

cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
CmdDel.Enabled = (nMode = LoadMode And bEditRecord)
cmdSave.Enabled = bEditRecord
cmdInform.Enabled = (nMode = LoadMode)


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
xCode.text = Newflag("MEETING", "code", con)
xcode_zero.Caption = ""
xdesca.text = ""
xDate2.text = ""
xDate.text = ""
cmdYear.Tag = sSeason
cmdYear.Caption = retFlag(aSeason, "desca")
cmdYearPaid.Caption = "«Œ «— «·„Ê”„"
cmdYearPaid.Tag = ""
xClosed.Value = 0
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xCode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
    myUndo
ElseIf ActiveControl.Name = cmdYear.Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
ElseIf ActiveControl.Name = cmdYearPaid.Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Sub myload()
xCode.text = CardTable!code & ""
xcode_zero.Caption = CardTable!CODE_ZERO & ""
xdesca.text = CardTable!Desca & ""
xDate.text = myFormat_p(CardTable!Date)
xDate2.text = myFormat_p(CardTable!Date2)
cmdYear.Tag = CardTable!SEASON & ""
cmdYear.Caption = GetField("SELECT DESCA FROM YEARS_CODES WHERE CODE = " & addvalue(CardTable!SEASON & ""), con)
cmdYearPaid.Tag = CardTable!SEASON_PAID & ""
If cmdYearPaid.Tag = "" Then
    cmdYearPaid.Caption = "«Œ «— «·„Ê”„"
Else
    cmdYearPaid.Caption = GetField("SELECT DESCA FROM YEARS_CODES WHERE CODE = " & addvalue(CardTable!SEASON_PAID & ""), con) & ""
End If
bCheck = True
xClosed.Value = IIf(CardTable!Closed, 1, 0)
bCheck = False
Handlecontrols LoadMode
End Sub
Private Function MyReplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xdesca.text))
aInsert = AddFlag(aInsert, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[DATE2]", addDate(xDate2.text))
aInsert = AddFlag(aInsert, "[SEASON]", addvalue(cmdYear.Tag))
aInsert = AddFlag(aInsert, "[SEASON_PAID]", addvalue(cmdYearPaid.Tag))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.text))
    con.Execute addInsert(aInsert, "MEETING")
Else
    con.Execute addUpdate(aInsert, "MEETING", "CODE = " & xCode.text)
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
xdesca.SetFocus
End Sub

Private Sub cmdAddMember_Click()
Me.MousePointer = 11
If AddMember2 Then
    Me.MousePointer = 0
    Inform "‰„ «÷«›… «⁄÷«¡«·Ã„⁄Ì… »‰Ã«Õ"
Else
    Me.MousePointer = 0
    Inform "·„   „ «÷«›… «⁄÷«¡«·Ã„⁄Ì… »‰Ã«Õ"
End If
End Sub
Private Function AddMember2() As Boolean
If Not validAdd Then Exit Function

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear.Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "CODE1", Null)
aPrm = AddFlag(aPrm, "CODE2", Null)
aPrm = AddFlag(aPrm, "DATE1", Null)
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

prog1.Value = 0
Dim loctable As ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_meet_member]", con, adStoredProc, aPrm, 300)

nRecordcount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True

con.BeginTrans
On Error GoTo myerror
con.Execute "delete from mem_Meeting where code_meeting = " & xCode.text
Do Until loctable.EOF
    i = i + 1
    prog1.Value = IIf(Round(i / (nRecordcount), 2) > 1, 1, Round(i / (nRecordcount), 2)) * 100
    con.Execute "insert into mem_meeting(code_meeting,CODE,CODE_REL,ID) values(" & xCode.text & "," & loctable!code & "," & loctable!relation & "," & addstring(retMeetCode(xCode.text, loctable!code, loctable!relation)) & ")"
    loctable.MoveNext
Loop
con.CommitTrans
prog1.Visible = False
AddMember2 = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdDel_Click()
Dim nDelete As Long
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From MEETING Where code = " & xCode.text, nDelete
    con.CommitTrans
    openCardTable xCode.text, "<="
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

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdInform_Click()
'MeetingLookup Me, oSearch
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
myUndo
End Sub
Private Sub CmdNext_Click()
openCardTable xcode_zero.Caption, ">"
If CardTable.EOF Then openCardTable xCode.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xcode_zero.Caption, "<"
If CardTable.EOF Then openCardTable xCode.text
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
Private Sub CmdUndo_Click()
myUndo
End Sub

Private Sub cmdYear_Click()
Years_LookupAll Me, oSearchYear
End Sub
Private Sub cmdYearPaid_Click()
Years_LookupAll Me, oSearchYear, , cmdYearPaid.Tag <> ""
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
bedit = True
myUndo
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.text) Then
     If xCode.Tag = LoadMode Then
        mydefine
    Else
        xCode.text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
        If CardTable!code = xCode.text Then
            Exit Sub
        End If
    End If
    
    openCardTable xCode.text
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
If Not ValidInt(xCode.text) Then
    If Not igMsg Then MsgBox "«·ﬂÊœ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not igMsg Then MsgBox " «—ÌŒ «·Ã„⁄Ì… €Ì— „”Ã·…", , systemName
    Exit Function
End If

If Not IsDate(xDate2.text) Then
    If Not igMsg Then MsgBox " «—ÌŒ «‰ Â«¡ «·”œ«œ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not ValidNum(cmdYear.Tag) Then
    MsgBox "„Ê”„ «·”œ«œ €Ì— „”Ã·", , systemName
    Exit Function
End If

If xdesca.text = "" Then
    If Not igMsg Then MsgBox "»Ì«‰ «·Ã„⁄Ì… €Ì— „”Ã·", , systemName
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
cString = "SELECT TOP 1 MEETING.* FROM MEETING"
If pSign = "=" Then
    If pCode <> "" Then cWhere = "CODE  " & pSign & addstring(pCode)
Else
    If pCode <> "" Then cWhere = "CODE_ZERO " & pSign & addstring(pCode)
End If

cFilter = ""
If sCode <> "" Then cFilter = "MEETING.CODE = " & addvalue(sCode)
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by MEETING.CODE_ZERO desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by MEETING.CODE_ZERO ASC"
End If

CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.text) Then
    openCardTable xCode.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
On Error Resume Next
xdesca.SetFocus
Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If Trim(pCode) <> "" Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE_ZERO <= " & MyParn(pCode) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM MEETING " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If Trim(pCode) <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub xCurrent_mem_Click()
If Not bCheck Then
    If xCurrent_mem.Value = 1 Then
        xNo.text = mRound(GetField("select top 1 no from MEETING where current_mem = 1 order by no desc", con) + 1)
    Else
        xNo.text = ""
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
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Function AddMember() As Boolean
If Not validAdd Then Exit Function
Dim loctable As ADODB.Recordset
Dim cString As String, cWhere As String, cWhere2 As String, i As Long
cString = "SELECT FILE1_10.CODE" & _
           " From File1_10 INNER JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE  " & _
           " AND FILE6_20H.doc_no = dbo.f_meeting_doc(FILE1_10.CODE," & cmdYear.Tag & ")"

cWhere = "FILE6_20H.DATE <= " & DateSq(xDate2.text)
cWhere = cWhere & Tr(cWhere) & "FILE1_10.Date_Begin <= " & DateSq(DateAdd("yyyy", -1, xDate.text))
If IsDate(xDate2.text) Then
    cWhere = cWhere & Tr(cWhere) & " FILE6_20H.DATE <= " & DateSq(xDate2.text)
End If
cWhere2 = cWhere & " AND " & "FILE1_10.DIED = 0"
If cWhere2 <> "" Then cString = cString & " WHERE " & cWhere2


Set loctable = New ADODB.Recordset
Set loctable = myCmd(cString, con)
'loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
con.BeginTrans
On Error GoTo myerror
con.Execute "delete from mem_Meeting where type = 0"

prog1.Value = 0
prog1.Visible = True
Do Until loctable.EOF
    i = i + 1
    prog1.Value = IIf(Round(i / (nRecordcount), 2) > 1, 1, Round(i / (nRecordcount), 2)) * 100
    If IsEmpty(GetField("select top 1 id from mem_meeting where id = " & MyParn(retMeetCode(xCode.text, loctable!code, "0")), con)) Then
        con.Execute "insert into mem_meeting(code_meeting,CODE,CODE_REL,ID) values(" & xCode.text & "," & loctable!code & ",0," & addstring(retMeetCode(xCode.text, loctable!code, "0")) & ")"
    End If
    loctable.MoveNext
Loop

prog1.Visible = False
cString = "SELECT FILE1_11.Member,file1_11.code" & _
          " from file1_11 inner join file1_10 on file1_11.member = file1_10.code INNER JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE" & _
           " AND FILE6_20H.doc_no = dbo.f_meeting_doc(FILE1_10.CODE," & cmdYear.Tag & ")"
cWhere = cWhere & turn(cWhere, " AND ") & "FILE1_11.RELATION  = 1"
cWhere = cWhere & " AND " & "FILE1_11.Date_Begin <= " & DateSq(DateAdd("yyyy", -1, xDate.text))
If cWhere <> "" Then cString = cString & " WHERE " & cWhere


Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
nRecordcount = loctable.RecordCount
On Error GoTo myerror
prog1.Value = 0
prog1.Visible = True
i = 0
Do Until loctable.EOF
    i = i + 1
    prog1.Value = IIf(Round(i / (nRecordcount), 2) > 1, 1, Round(i / (nRecordcount), 2)) * 100
    If IsEmpty(GetField("select top 1 id from mem_meeting where id = " & MyParn(retMeetCode(xCode.text, loctable!member, loctable!code)), con)) Then
        con.Execute "insert into mem_meeting(code_meeting,CODE,CODE_REL,ID) values(" & xCode.text & "," & loctable!member & "," & loctable!code & "," & addstring(retMeetCode(xCode.text, loctable!member, loctable!code)) & ")"
    End If
    loctable.MoveNext
Loop
con.CommitTrans
prog1.Visible = False
AddMember = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function retMeetCode(pMeet As String, pMember As String, pRel As String) As String
retMeetCode = RetZero(pMeet, 2) & "-" & RetZero(pMember, 10) & "-" & RetZero(pRel, 2)
End Function
Private Function validAdd() As Boolean
If Not ValidInt(xCode.text) Then
    If Not igMsg Then MsgBox "«·ﬂÊœ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not igMsg Then MsgBox " «—ÌŒ «·Ã„⁄Ì… €Ì— „”Ã·…", , systemName
    Exit Function
End If

If Not IsDate(xDate2.text) Then
    If Not igMsg Then MsgBox " «—ÌŒ «‰ Â«¡ «·”œ«œ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not ValidNum(cmdYear.Tag) Then
    MsgBox "„Ê”„ «·”œ«œ €Ì— „”Ã·", , systemName
    Exit Function
End If
validAdd = True
End Function

