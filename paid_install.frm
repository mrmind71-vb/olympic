VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paid_installfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ì’«·«  ”œ«œ"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16980
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
   ScaleHeight     =   9555
   ScaleWidth      =   16980
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame20 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   -45
      Width           =   3750
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«€·«ﬁ „” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   1365
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   420
         Left            =   90
         TabIndex        =   48
         Top             =   225
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«€·«ﬁ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
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
      Height          =   1230
      Index           =   0
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1575
      Width           =   1950
      Begin Threed.SSCommand cmdAddItems 
         Height          =   1050
         Left            =   45
         TabIndex        =   37
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   1852
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "«÷«›… «·ﬁ”ÿ"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   690
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   6180
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4995
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":23CB
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3735
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":4474
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":647C
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2475
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":8433
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":ABCF
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":D063
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":F386
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid_install.frx":116FC
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2130
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   675
      Width           =   7620
      Begin VB.TextBox xcard_price 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4500
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1665
         Width           =   825
      End
      Begin VB.TextBox xCard_count 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5355
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1665
         Width           =   780
      End
      Begin VB.TextBox xOther 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "N"
         Top             =   1710
         Width           =   1320
      End
      Begin VB.TextBox xForm_No2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         Locked          =   -1  'True
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   945
         Width           =   1635
      End
      Begin VB.TextBox xCard_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2835
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "N"
         Top             =   1665
         Width           =   1635
      End
      Begin VB.TextBox xForm_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4500
         Locked          =   -1  'True
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   945
         Width           =   1635
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4500
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "N"
         Top             =   1305
         Width           =   1635
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4500
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1635
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4500
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   1635
      End
      Begin Threed.SSCommand cmdData 
         Height          =   375
         Left            =   90
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "»Ì«‰«  «·⁄÷Ê"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":1387F
      End
      Begin Threed.SSCommand cmdFixCard 
         Height          =   330
         Left            =   2160
         TabIndex        =   59
         Top             =   1665
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
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
         Caption         =   "..."
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   " »—⁄"
         Height          =   285
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1755
         Width           =   405
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ«—‰ÌÂ« "
         Height          =   240
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1710
         Width           =   1125
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ﬁ”Ì„…"
         Height          =   240
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1350
         Width           =   1125
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1395
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1305
         Width           =   3075
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   510
      End
   End
   Begin VB.Frame FRAME_CUR 
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
      Height          =   1230
      Index           =   4
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1575
      Width           =   1770
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   900
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
         Picture         =   "paid_install.frx":15B6C
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":18491
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   675
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "paid_install.frx":1ACE5
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid_install.frx":1CE45
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   1890
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   2
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4455
      Top             =   8190
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   2115
      Top             =   8280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   38
      Top             =   9090
      Width           =   16980
      _ExtentX        =   29951
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   39
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
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
         Left            =   4095
         TabIndex        =   40
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
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
         Left            =   8100
         TabIndex        =   41
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
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
         Index           =   3
         Left            =   12150
         TabIndex        =   42
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
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
         Index           =   4
         Left            =   16155
         TabIndex        =   43
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4335
      Left            =   90
      TabIndex        =   10
      Top             =   2835
      Width           =   16575
      _cx             =   29236
      _cy             =   7646
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
      WordWrap        =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   135
      Top             =   1260
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Frame FRAME10 
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
      Height          =   1365
      Left            =   10665
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   7245
      Width           =   6045
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„’«—Ì›"
         Height          =   240
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label xCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label xInterest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›«∆œ…"
         Height          =   240
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… „÷«›…"
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label xTotal_Tax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label xTotal_Value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… «·«‘ —«ﬂ"
         Height          =   240
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ï"
         Height          =   285
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7155
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   31
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
         Picture         =   "paid_install.frx":1F132
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid_install.frx":212D9
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   32
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
         Picture         =   "paid_install.frx":23320
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid_install.frx":2540B
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   33
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
         Picture         =   "paid_install.frx":27405
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid_install.frx":29516
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   34
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
         Picture         =   "paid_install.frx":2B510
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid_install.frx":2D734
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   49
      Top             =   8895
      Visible         =   0   'False
      Width           =   16980
      _ExtentX        =   29951
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   7155
      Width           =   6000
      Begin VB.CheckBox chkYear 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·”‰… «·Õ«·Ì…"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   225
         Width           =   1365
      End
      Begin VB.CheckBox chkMonth 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘Â— «·Õ«·Ì"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   1365
      End
      Begin VB.CheckBox chkDay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·ÌÊ„"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.Label xOTHER_VALUE_TAX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   945
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   1665
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "paid_installfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, sCode As String, bNew As Boolean, bMem As Boolean
Public myForm As Form, bFawry As Boolean
Dim cList As String, cFilter As String, bClosed As Boolean
Dim CardTable As ADODB.Recordset, loctable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim oSearchDoc As New Search3, oSearchMember As New Search, oSearchInstall As New Search_empty, oSearchRel As New Search3, oSearchYearChange As New Search_empty
Dim bEditRecord As Boolean, bAct As Boolean, aPen As Variant
Dim DocTitle As String
Dim bIgChange As Boolean
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, cdef_Box As String, bNewSave As Boolean
Dim formMode
Dim bCheck As Boolean
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant, i As Integer
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.text))
aInsert = AddFlag(aInsert, "FORM_NO", addvalue(xForm_No.text))
aInsert = AddFlag(aInsert, "FORM_NO2", addvalue(xForm_No2.text))
aInsert = AddFlag(aInsert, "CARD_COUNT", mRound(xCard_count.text))
aInsert = AddFlag(aInsert, "CARD_PRICE", mRound(xcard_price.text))
aInsert = AddFlag(aInsert, "CARD_VALUE", mRound(xCard_value.text))
aInsert = AddFlag(aInsert, "OTHER", mRound(xOther.text))

aInsert = AddFlag(aInsert, IIf(xdoc_no.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName & " [" & GetComputerName & "]"))
aInsert = AddFlag(aInsert, IIf(xdoc_no.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")

con.BeginTrans
On Error GoTo myerror
If xdoc_no.Tag = DefineMode Then
    xdoc_no.text = Newflag("FILE6_30H", "DOC_NO")
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xdoc_no.text))
    con.Execute addInsert(aInsert, "FILE6_30H")
Else
    con.Execute addUpdate(aInsert, "FILE6_30H", "doc_no = " & addstring(xdoc_no.text))
End If
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xdoc_no.text))
        aInsert = AddFlag(aInsert, "VALUE", Val(.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "TAX_RATE", Val(.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "TAX_DIFF", Val(.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "RATE_INTEREST", .TextMatrix(i, 8))
        aInsert = AddFlag(aInsert, "LATE_ID", addvalue(.TextMatrix(i, .Cols - 2)))
        If .TextMatrix(i, .Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE6_30")
        Else
            con.Execute addUpdate(aInsert, "FILE6_30", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
    Unload oSearchMember
ElseIf ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, grid1.Cols - 2) = oSearchInstall.grid1.TextMatrix(oSearchInstall.grid1.Row, oSearchInstall.grid1.Cols - 1)
        grid1.TextMatrix(grid1.Row, 4) = oSearchInstall.grid1.TextMatrix(oSearchInstall.grid1.Row, 5)
        grid1.TextMatrix(grid1.Row, 5) = oSearchInstall.grid1.TextMatrix(oSearchInstall.grid1.Row, 3)
        'GrdDesc oSearchInstall.grid1.TextMatrix(oSearchInstall.grid1.Row, 0), grid1.Row
        Grid1_AfterEdit grid1.Row, grid1.Col
        Unload oSearchInstall
        CellPos 13, grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    xdoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
End If
End Sub
Private Sub cmd_closed_Click()
con.BeginTrans
On Error GoTo myerror
con.Execute " update " & cFileHeader & " set CLOSED = " & IIf(xClosed.Value = 1, "0", "1") & " WHERE doc_no = " & MyParn(xdoc_no.text)
con.CommitTrans
Err.Clear
'openCardTable
myUndo
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub

Private Sub chkDay_Click()
If Not bCheck Then myUndo
End Sub

Private Sub chkMonth_Click()
If Not bCheck Then myUndo
End Sub

Private Sub chkYear_Click()
If Not bCheck Then myUndo
End Sub

Private Sub cmdAddItems_Click()
myAdditems
End Sub
Private Function myAdditems() As Boolean
Dim nYears As Long, nFirstYear As Integer, aRet As Variant
If Not ValidNum(xCode.text) Then
    MsgBox "ﬂÊœ «·⁄÷Ê €Ì— ’ÕÌÕ"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If

If bNew Then
    aRet = DocSameDay_i(xCode.text, xDate.text, con)
    If Not IsEmpty(aRet) Then
        MsgBox "„” ‰œ »‰›” ‰Ê⁄ «·„ÿ«·»… »‰›” «·ÌÊ„ —ﬁ„ " & aRet
        xdoc_no.text = aRet
        xDoc_No_LostFocus
        Exit Function
    End If
End If

If bFawry Then
    Dim nTotalMax As Double
    nTotalMax = mRound(GetField("[dbo].fawry_acount_install(file2_10.code)"))
    myAddInstall (nTotalMax + 0.25)
Else
    myAddInstall
End If
End Function
Private Function AddMemberData(aMember As Variant, Index As Variant) As Boolean
Dim nAge As Integer, nGender As Integer
If IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
   nAge = Age(myFormat(retFlag(aMember, "DATE_BIRTH")), myFormat(xDate.text)) - Index
Else
   nAge = 1
End If

If Val(loctable!Age1 & "") > nAge And Val(loctable!Age1) <> 0 Then Exit Function
If Val(loctable!Age2 & "") < nAge And Val(loctable!Age2 & "") <> 0 Then Exit Function
If (Not IsNull(loctable!GENDER)) Then
    nGender = TurnValue(retFlag(aMember, "Gender", 1), Null, 1)
    If nGender <> loctable!GENDER Then Exit Function
End If
AddMemberData = True
End Function
Private Sub cmdClosePeriod_Click()
closefrm.sFile = "FILE6_30H"
closefrm.Show 1
myUndo
End Sub

Private Sub cmdData_Click()
Dim oMember As New memberfrm
If ValidNum(xCode.text) Then
    oMember.sCode = xCode.text
    oMember.Show
End If
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE6_30 where Doc_No = " & xdoc_no.text
    con.Execute "Delete From FILE6_30H where Doc_No = " & xdoc_no.text
    con.CommitTrans
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
    
    'openCardTable xdoc_no_zero.Caption, "<="
    openCardTable xdoc_no.text, "<="
    If CardTable.EOF Then openCardTable , ">"
    If CardTable.EOF Then
        mydefine
    Else
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Function retAll(aMember As Variant) As Integer
retAll = IIf(retFlag(aMember, "Died", False), 0, 1)
Dim cString As String
cString = "SELECT SUM(1) FROM FILE1_11"
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.text
retAll = retAll + Val(GetField(cString, con) & "")
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_30H.DOC_NO, FILE6_30H.FORM_NO,CONVERT(VARCHAR(10),FILE6_30H.DATE,111), FILE2_10.DESCA,FILE6_30H.TOTAL" & _
          "  FROM  FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_30H.DATE,FILE6_30H.Doc_No"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·«” „«—…- «—ÌŒ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE2_10.Desca%% or **FILE6_30H.FORM_NO**" & _
                  " OR ##FILE6_30H.Date##)"

listarray(1, 0) = "ﬂÊœ «·⁄÷Ê"
listarray(1, 1) = "(**FILE6_30H.CODE**)"

listarray(2, 0) = "—ﬁ„ «·„” ‰œ"
listarray(2, 1) = "(**FILE6_30H.DOC_NO**)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1300

GrdArray(1, 0) = "—ﬁ„ «·«Ì’«·"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(2, 1) = 1400

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 5000

GrdArray(4, 0) = "«·≈Ã„«·Ì"
GrdArray(4, 1) = 1100

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„ «·„ÿ«·»« "
oSearchDoc.Show 1
End Sub

Private Sub cmdFixCard_Click()
Dim loctable As New ADODB.Recordset
Dim cString As String
bIgChange = True
xCard_count.text = myField("Select Count(*)  as CountOf from file2_11 where member = " & xCode.text, "CountOf", con, , 0) + 1
xcard_price.text = myField("Select top 1 rate2  from address", "rate2", con, , 0)
CalcTotals
bIgChange = False
End Sub

Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdNext_Click()
'openCardTable xdoc_no_zero.Caption, ">"
openCardTable xdoc_no.text, ">"
If CardTable.EOF Then openCardTable xdoc_no.text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
'openCardTable xdoc_no_zero.Caption, "<"
openCardTable xdoc_no.text, "<"
If CardTable.EOF Then openCardTable xdoc_no.text, "="
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
Private Sub cmdNewInv_Click()
mydefine
On Error Resume Next
xCode.SetFocus
Err.Clear
End Sub

Private Sub cmdPrint_Click()
doprint
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
'openCardTable
If sDoc_no <> "" Or (Not myForm Is Nothing) Or bMem Then
 '   myForm.myrefresh
    Unload Me
    Exit Sub
End If
myUndo
End Sub
Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdYearChange_Click()
If xType.BoundText = "2" And xdoc_no.Tag = LoadMode Then
    Years_LookupAll Me, oSearchYearChange
ElseIf xdoc_no.Tag = DefineMode Then
    Years_LookupAll Me, oSearchYearChange, "change_season"
End If
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If sCode <> "" Then
        xCode.text = sCode
        xCode_LostFocus
        If bNew Then
            mydefine
            cmdAddItems.SetFocus
            myAdditems
        End If
    Else
        If xdoc_no.Tag = LoadMode Then grid1.SetFocus Else xCode.SetFocus
    End If
    Err.Clear
    bNewSave = bNew
    bNew = False
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And ActiveControl.Name <> xCode.Name Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
bSlow = True
openCon con

Set grid1.DataSource = data1

bCheck = True
LoadText Me
bCheck = False

openCardTable
If bNew And sCode <> "" Then
    If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE6_30H.CODE = " & addvalue(sCode)
    mydefine
ElseIf sDoc_no <> "" Then
    myUndo
Else
    mydefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(chkYear.Name, chkMonth.Name, chkYear.Name)
CardTable.Close
Set CardTable = Nothing
closeCon con
If (Not myForm Is Nothing) Then
    myForm.myrefresh
End If
Set paid_installfrm = Nothing
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not MYVALID(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myLoadGrd
    If Row < grid1.rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If

With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.rows - 1 Then
    MyAddItem
End If

CalcTotals

On Error GoTo myerror
If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    End If
    CalcTotals
Else
    myLoadGrd
End If
End With
Exit Sub
myerror:
myLoadGrd
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    InstallLookup Me, oSearchInstall, "FILE6_21.CODE = " & addvalue(xCode.text), xdoc_no.text
ElseIf Col = 9 Then
    Dim oInterest As New ShowInterestfrm
    oInterest.pCode = xCode.text
    oInterest.pDate = xDate.text
    oInterest.pId = grid1.TextMatrix(Row, grid1.Cols - 2)
    oInterest.Show 1
End If
End Sub
Private Sub grid1_EnterCell()
With grid1
    If Not bEditRecord Then
        .Editable = flexEDNone
    ElseIf ((.Col = 0 And .TextMatrix(.Row, .Cols - 1) = "" And ValidNum(xCode.text)) Or .Col = 4 Or .Col = 7 Or .Col = 8 Or .Col = 9) Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

'If Not bIgMsg Then
'    If grid1.rows < 3 Then
'        MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
'        Exit Function
'    End If
'End If

With grid1
For i = 1 To .rows - 2
'    If .TextMatrix(i, 1) = "" Then
'        .Select i, 0, i, grid1.Cols - 1
'        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
'        Exit Function
'    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
Dim i As Integer
xdoc_no.text = CardTable!doc_no
'xdoc_no_zero.Caption = CardTable!doc_no_zero & ""
xForm_No.text = CardTable!FORM_NO & ""
xForm_No2.text = CardTable!FORM_NO2 & ""
xDate.text = myFormat_p(CardTable!Date)
bIgChange = True
xCard_count.text = Myvalue(CardTable!card_Count)
xcard_price.text = Myvalue(CardTable!card_price)
bIgChange = False
xCard_value.text = Myvalue(CardTable!CARD_VALUE)
xOther.text = Myvalue(CardTable!OTHER)
xOTHER_VALUE_TAX.Caption = Myvalue(CardTable!other_value_tax)
xCode.text = CardTable!code & ""
LoadMember

xTotal.Caption = Myvalue(CardTable!total)

'xSeason.Caption = CardTable!SEASON
bClosed = True
xClosed.Value = IIf(CardTable!Closed, 1, 0)
bClosed = False

panel1(1).Caption = CardTable!UserName & ""
panel1(2).Caption = CardTable!UserName2 & ""
panel1(3).Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")
panel1(4).Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")

Handlecontrols LoadMode

myLoadGrd
On Error Resume Next
grid1.SetFocus
CellPos 13, grid1.rows - 2, grid1.Cols - 1
Err.Clear
End Sub
Private Function myLoadGrd(Optional bRefresh As Boolean = True) As Boolean
Dim cString As String
Dim afield(14)

afield(0) = "dbo.f_serial(FILE6_21.CODE,FILE6_21.DATE_DUE)"
afield(1) = "convert(varchar(10),FILE6_21.DATE_DUE,111)"
afield(2) = "FILE6_21.VALUE"
afield(3) = "FILE6_21.VALUE - dbo.f_install_paid(FILE6_21.ID)"
afield(4) = "FILE6_30.VALUE"
afield(5) = "FILE6_30.TAX_RATE"
afield(6) = "FILE6_30.TAX"
afield(7) = "FILE6_30.TAX_DIFF"
afield(8) = "FILE6_30.RATE_INTEREST"
afield(9) = "FILE6_30.INTEREST"
afield(9) = "FILE6_30.INTEREST"
afield(10) = "FILE6_30.TOTAL"
afield(11) = "FILE6_30.NOTES"
afield(12) = "FILE6_30.LATE_ID"
afield(13) = "FILE6_30.FILE6_30.[ID]"

cString = "SELECT dbo.f_serial(FILE6_21.CODE,FILE6_21.DATE_DUE)," & _
          "convert(varchar(10),FILE6_21.DATE_DUE,111)," & _
          "FILE6_21.VALUE," & _
          " FILE6_21.VALUE - dbo.f_install_paid(FILE6_21.ID)," & _
          " FILE6_30.VALUE," & _
          "FILE6_30.TAX_RATE," & _
          " FILE6_30.TAX," & _
          " FILE6_30.TAX_DIFF," & _
          " FILE6_30.RATE_INTEREST," & _
          " FILE6_30.INTEREST," & _
          " FILE6_30.CHARGE," & _
          " FILE6_30.TOTAL," & _
          " FILE6_30.NOTES," & _
          " FILE6_30.LATE_ID," & _
          " FILE6_30.[ID] " & _
          " From FILE6_30 LEFT JOIN FILE6_21 ON FILE6_30.LATE_ID = FILE6_21.ID"
cString = cString & turn(cString) & "FILE6_30.DOC_NO = " & addvalue(xdoc_no.text)
Set data1.Recordset = myRecordSet(cString, con)
MyAddItem
CalcTotals
Fixgrd
End Function
Private Sub mydefine()
Dim i As Integer, aRet As Variant
xdoc_no.text = Newflag("FILE6_30H", "DOC_NO")

'xdoc_no_zero.Caption = ""
'xForm_no.Text = Newflag(cFileHeader, "FORM_NO", con, "SEASON = " & sSeason)
xForm_No.text = ""
xForm_No2.text = ""

xCard_value.text = ""
xcard_price.text = ""
xCard_count.text = ""

bIgChange = True
xOther.text = ""
xOTHER_VALUE_TAX.Caption = ""
bIgChange = False

'xType.BoundText = "1"

bClosed = True
xClosed.Value = 0
bClosed = False


xDate.text = myFormat_p(Date)


xCode.text = sCode
xCodeDesca.Caption = ""

panel1(1).Caption = ""
panel1(2).Caption = ""
panel1(3).Caption = ""
panel1(4).Caption = ""



Fixgrd
grid1.rows = 1
MyAddItem

Handlecontrols DefineMode
CalcTotals
On Error Resume Next
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bedit And xClosed.Value = 0 And (xForm_No.text = "")
If sDoc_no <> "" Then bEditRecord = bEditRecord And xdoc_no.text = sDoc_no
cmdNewInv.Enabled = (Not bNew) And (Not bMem) And nMode = LoadMode And bedit
cmdAddItems.Enabled = bEditRecord
'cmdYearChange.Enabled = nMode = LoadMode And xType.BoundText = "2"
'cmdFilter.Visible = cmdFilter.Tag <> ""
'cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdFixCard.Enabled = bEditRecord
cmdSave.Enabled = bEditRecord
CmdDel.Enabled = nMode = LoadMode And bEditRecord
'xdate.Enabled = nMode = Mode
'xDate.Locked = True
'xForm_no.Locked = True
xCode.Locked = nMode = LoadMode

aRecords = retRecords(xdoc_no.text)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If
cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xClosed.Enabled = nMode = LoadMode
xClosed.Enabled = xClosed.Enabled And (bopt1 Or xClosed.Value = 0)

xdoc_no.Enabled = (nMode = DefineMode)
xdoc_no.Tag = nMode
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then Exit Sub
With grid1
    If KeyCode = 13 Then
        CellPos KeyCode, .Row, .Col
    ElseIf KeyCode = 46 And .Row <> .rows - 1 And .rows > 2 And bEditRecord Then
        If MsgBox("Õ–› øø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            myDeleteRow grid1.Row
            grid1_EnterCell
        End If
    End If
End With
End Sub
Private Function myDeleteRow(Row As Long) As Boolean
With grid1
con.BeginTrans
On Error GoTo myerror
If .TextMatrix(Row, .Cols - 1) <> "" Then
    con.Execute "Delete from FILE6_30 where ID = " & .TextMatrix(Row, .Cols - 1)
End If
con.CommitTrans
myRemove Row
End With
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If (grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    ElseIf Not ValidInt(grid1.EditText) Then
        MsgBox "«·ﬂÊœ €Ì— ”·Ì„"
        Cancel = True
    Else
        If Not GrdDesc(grid1.EditText, Row) Then
           MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ «Ê ·« Ì’·Õ"
           Cancel = True
        End If
    End If
End If
End Sub

Private Sub grid2_Click()
If myLoadGrd(False) Then
    CellPos 13, grid1.rows - 2, grid1.Cols - 1
End If
End Sub
Private Sub Grid2_EnterCell()
If myLoadGrd(False) Then
    CellPos 13, grid1.rows - 2, grid1.Cols - 1
End If
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub xCard_count_Change()
If Not bIgChange Then
    CalcTotals
End If
End Sub

Private Sub xcard_price_Change()
If Not bIgChange Then
    CalcTotals
End If
End Sub

Private Sub xClosed_Click()
If Not bClosed Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "UPDATE FILE6_30H SET FILE6_30H.CLOSED = " & xClosed.Value & " WHERE DOC_NO = " & addvalue(xdoc_no.text)
    con.CommitTrans
    Inform IIf(xClosed.Value = 1, " „ «€·«ﬁ «·„” ‰œ", " „ › Õ «·„” ‰œ")
    myUndo
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupInstall Me, oSearchMember
ElseIf KeyCode = 13 And cmdAddItems.Enabled Then
    cmdAddItems_Click
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
LoadMember
End Sub
Private Sub LoadMember()
xCodeDesca.Caption = ""

If Not ValidNum(xCode.text) Then Exit Sub
Dim aMember As Variant
aMember = Member_Load_install(xCode.text, , con)
xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
End Sub
Private Sub xCurrent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'openCardTable
If Not bCheck Then
    myUndo
End If
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
If Not ValidNum(xdoc_no.text) Then
     If xdoc_no.Tag = LoadMode Then
        mydefine
    Else
        xdoc_no.text = Newflag("FILE6_30H", "DOC_NO", con)
    End If
Else
    If xdoc_no.Tag = LoadMode Then
        If (Not (CardTable.EOF)) Then
            If CardTable!doc_no = xdoc_no.text Then
                Exit Sub
            End If
        End If
    End If
    
    openCardTable xdoc_no.text
    If Not CardTable.EOF Then
        myload
    ElseIf xdoc_no.Tag = LoadMode Then
        mydefine
    Else
        xdoc_no.text = Newflag("FILE6_30H", "DOC_NO", con)
    End If
End If
End Sub
Private Sub xForm_no_GotFocus()
myGotFocus xForm_No
End Sub
Private Sub xForm_no_LostFocus()
myLostFocus xForm_No
End Sub
Private Sub xForm_no2_GotFocus()
myGotFocus xForm_No2
End Sub
Private Sub xForm_no2_LostFocus()
myLostFocus xForm_No2
End Sub
Private Sub xcard_value_GotFocus()
myGotFocus xCard_value
End Sub
Private Sub xcard_value_LostFocus()
myLostFocus xCard_value
End Sub
Private Function CalcTotals(Optional bOverRide As Boolean = False)
Dim nTotalRow As Double, nTotalItem As Double, Row As Integer
Dim Rate_Tax As Double, nTax As Double, nRate_Discount As Double, nTotalTax As Double, nRate_Late As Double, nLate As Double
Dim nChrage As Double
Dim nInterest As Double
With grid1

Dim afield(12)
afield(0) = "dbo.f_serial(FILE6_21.CODE,FILE6_21.DATE_DUE)"
afield(1) = "convert(varchar(10),FILE6_21.DATE_DUE,111)"
afield(2) = "FILE6_21.VALUE"
afield(3) = "FILE6_21.VALUE - dbo.f_install_paid(FILE6_21.ID)"
afield(4) = "FILE6_30.VALUE"
afield(5) = "FILE6_30.TAX_RATE"
afield(6) = "FILE6_30.TAX"
afield(7) = "FILE6_30.RATE_INTEREST"
afield(8) = "FILE6_30.INTEREST"
afield(9) = "FILE6_30.TOTAL"
afield(10) = "FILE6_30.NOTES"
afield(11) = "FILE6_30.LATE_ID"
afield(12) = "FILE6_30.FILE6_30.[ID]"

xCard_value.text = Myvalue(mRound(xCard_count.text) * mRound(xcard_price.text))

For Row = 1 To grid1.rows - 2
    nTotalRow = mRound(grid1.ValueMatrix(Row, 4))
    nTotalItem = nTotalItem + nTotalRow
    
    nRate_Tax = mRound(.ValueMatrix(Row, 5)) / 100
    
    nTax = mRound(mRound(nTotalRow) * mRound(nRate_Tax))
    
    grid1.TextMatrix(Row, 6) = mRound(mRound(nTotalRow) * mRound(nRate_Tax))
    
    nTax = nTax + mRound(.ValueMatrix(Row, 7))
    
    nTotalTax = nTax + nTotalTax
    
    .TextMatrix(Row, 9) = mRound(.ValueMatrix(Row, 4) * mRound((.ValueMatrix(Row, 8) / 100), 4))
    
    nInterest = nInterest + mRound(.ValueMatrix(Row, 9), 2)
    
    ncharge = ncharge + mRound(.ValueMatrix(Row, 10), 2)
    
    nTotalRow = nTotalRow + nTax + mRound(.ValueMatrix(Row, 9), 2) + mRound(.ValueMatrix(Row, 10), 2)
    
    nTotal = nTotal + nTotalRow
    
   .TextMatrix(Row, 10 + 1) = nTotalRow
Next

nTotalTax = nTotalTax + Val(xOTHER_VALUE_TAX.Caption)
xTotal_Value.Caption = Myvalue(nTotalItem)
xTotal_Tax.Caption = Myvalue(nTotalTax)
xInterest.Caption = Myvalue(nInterest)
xCharge.Caption = Myvalue(ncharge)

xTotal.Caption = Myvalue(mRound(nTotal) + Val(xOTHER_VALUE_TAX.Caption) + mRound(xCard_value.text) + mRound(xOther.text))
End With
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
'If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
.RowHeight(0) = 700

.FormatString = "—ﬁ„ «·ﬁ”ÿ|" & " «—ÌŒ «·«” Õﬁ«ﬁ|" & "ﬁÌ„… «·ﬁ”ÿ|" & "€Ì— „”œœ|" & "«·”œ«œ|" & "‰”»… ﬁÌ„… „÷«›…|" & "ﬁÌ„… „÷«›…|" & "›—ﬁ ﬁÌ„… „÷«›…|" & "‰”»… «·›«∆œ…|" & "›«∆œ…|" & "„’«—Ì›|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|" & "ﬂÊœ|"
.ColWidth(0) = 800
.ColWidth(1) = 1400
.ColWidth(2) = 1200
.ColWidth(3) = 1200
.ColWidth(4) = 1200
.ColWidth(5) = 1200
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(10 + 1) = 1200
.ColWidth(11 + 1) = 1500
.ColDataType(3) = flexDTDecimal
.ColDataType(4) = flexDTDecimal
.ColDataType(5) = flexDTDecimal
.ColDataType(6) = flexDTDecimal
.ColDataType(7) = flexDTDecimal
.ColDataType(8) = flexDTDecimal
.ColDataType(9) = flexDTDecimal
.ColDataType(10) = flexDTDecimal
.ColDataType(10 + 1) = flexDTDecimal
.ColDataType(11 + 1) = flexDTDecimal
.ColComboList(0) = "..."
.ColComboList(9) = "..."

'.ColHidden(4) = True
'.ColHidden(6) = True
'.ColHidden(7) = True
'.ColHidden(9) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
'.ColComboList(0) = cList
End With
End Sub
Private Sub fixgrd2()
With grid2
.FormatString = "«·”‰…|" & "«·ﬂÊœ"
.ColWidth(0) = 2000
'.ColWidth(1) = 1200
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Fixgrd3()
With grid3
'.FormatString = "«·”‰…|" & "«·≈Ã„«·Ì|" & "«·ﬂÊœ"
'.ColWidth(0) = 2000
'.ColWidth(1) = 1200
'.ColHidden(.Cols - 1) = True
'For i = 0 To .Cols - 1
'    .ColAlignment(i) = flexAlignRightCenter
'Next
End With
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String

Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1  FILE6_30H.* FROM FILE6_30H"

If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_No)
    'If pDoc_No <> "" Then cWhere = "DOC_NO_ZERO " & pSign & addstring(pDoc_No)
End If

'If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_No)

cFilter = ""

If chkDay.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "FILE6_30H.[DATE] = " & DateSq(Date)
If chkMonth.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_30H.[DATE]) = " & Year(Date) & " AND MONTH(FILE6_30H.DATE) = " & Month(Date)
If chkYear.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_30H.[DATE]) = " & Year(Date)


If sDoc_no <> "" Then cFilter = "FILE6_30H.DOC_NO = " & addvalue(sDoc_no)

If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
     cString = cString & " order by FILE6_30H.doc_no desc"
    'cString = cString & " order by FILE6_30H.doc_no_zero desc"
ElseIf pSign = ">=" Or pSign = ">" Then
     cString = cString & " order by FILE6_30H.doc_no ASC"
    'cString = cString & " order by FILE6_30H.doc_no_zero ASC"
End If
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Function retRecords(pDoc_No) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    'cString = "SELECT SUM(1) AS records,SUM(CASE WHEN doc_no_zero <= " & MyParn(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN doc_no <= " & addvalue(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM FILE6_30H " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If pDoc_No <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub myUndo()
'On Error GoTo myerror
Dim cString As String
If ValidNum(xdoc_no.text) Then
    openCardTable xdoc_no.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub MyAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
'If (Not validRow(grid1.Row)) And grid1.Row <> grid1.rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub
Private Function validRow(Row) As Boolean
With grid1
If .TextMatrix(Row, grid1.Cols - 2) = "" Then Exit Function
If mRound(.TextMatrix(Row, 4)) = 0 Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
With grid1
If Col < .Cols - 7 Then
    .Col = Col + 1 + IIf(Col = 0, 2, 0)
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 3)
    .ShowCell .Row, 0
Else
    .Select Row, Col
End If
End With
End Sub
Private Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For i = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(Row, i)) = "" Then
        NextEmpty = i
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Function GrdDesc(sId As String, Row As Long) As Boolean
'Dim sSection As String, aRet As Variant
'If ValidNum(sItem) Then
'
'    If ValidNum(xCode.Text) And RetYearCode Then
'        aRet = GetFields("SELECT TOP 1 DESCA,dbo.f_mem_item_value(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS VALUE,dbo.f_mem_item_discount(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS DISCOUNT,dbo.f_mem_item_tax(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS TAX   FROM file6_10 where ITEM = " & sItem)
'    Else
'        aRet = GetFields("SELECT TOP 1 DESCA,VALUE FROM file6_10 where ITEM = " & sItem)
'    End If
'    grid1.TextMatrix(Row, 1) = retFlag(aRet, "DESCA") & ""
'    grid1.TextMatrix(Row, 2) = retFlag(aRet, "VALUE") & ""
'    If grid1.TextMatrix(Row, 3) = "" Then grid1.TextMatrix(Row, 3) = 1
'    grid1.TextMatrix(Row, 4) = mRound(retFlag(aRet, "DISCOUNT"))
'    grid1.TextMatrix(Row, 5) = mRound(retFlag(aRet, "TAX"))
'End If
'GrdDesc = True
End Function
Private Function doprint()
If Not MYVALID Then Exit Function
Dim loctable As ADODB.Recordset, cString As String
Dim temptable As New ADODB.Recordset
cString = "SELECT V_FILE6_30.DOC_NO,TYPE_CODES.DESCA AS TYPE_DESCA,FILE6_30H.CODE,FILE6_30H.DATE," & _
          "FILE6_30H.TOTAL,FILE6_30H.TOTAL_VALUE,FILE6_30H.CHARGE,FILE6_30H.TOTAL_TAX,FILE2_10.DESCA AS DESCA_MEMBER,V_FILE6_30.DESCA," & _
          " V_FILE6_30.VALUE,V_FILE6_30.TOTAL AS TOTAL_ROW,V_FILE6_30.TAX + V_FILE6_30.TAX_DIFF AS TAX,V_FILE6_30.INTEREST,FILE6_30H.CARD_VALUE" & _
          " FROM FILE6_30H LEFT JOIN V_FILE6_30 ON FILE6_30H.DOC_NO = V_FILE6_30.DOC_NO" & _
          " INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE" & _
          " INNER JOIN TYPE_CODES ON FILE2_10.TYPE = TYPE_CODES.CODE"
cString = cString & turn(cString) & "FILE6_30H.DOC_NO = " & xdoc_no.text


Dim aTotal As Variant
'aTotal = GetFields("Select sum(V_FILE6_30.total) as total from V_FILE6_30 where doc_no = " & xDoc_No.Text)
Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim i As Long
With loctable
Do Until loctable.EOF
    temptable.AddNew
    temptable!str1 = ArbString(myFormat_p(loctable!Date))
    temptable!str2 = loctable!Desca_Member
    temptable!str3 = ArbString(loctable!code)
    temptable!str4 = "⁄÷ÊÌ… „ﬁ”ÿ…"
    temptable!str5 = ArbString(loctable!doc_no)
    temptable!val1 = mRound(loctable!Value)
    
    temptable!val2 = mRound(loctable!TAX)
    
    temptable!VAL3 = mRound(loctable!interest)
    temptable!val4 = mRound(loctable!total_row)
    temptable!str15 = ArbString(MyOnly(mRound(loctable!total)))
    
    temptable!val8 = mRound(loctable!total)
    'temptable!val9 = mRound(loctable!charge)
        
    'temptable!str11 = TurnValue(loctable!Item_Desca)
    'temptable!str12 = TurnValue(loctable!Desca)
    'temptable!str13 = TurnValue(loctable!notes)
    'temptable!str13 = TurnValue(loctable!notes)
    'temptable!str14 = TurnValue(loctable!user_name)
    temptable!str7 = loctable!Desca
    temptable!str21 = "≈Ì’«· ”œ«œ ﬁ”ÿ"
    'temptable!VAL1 = Val(loctable! & "")
    'temptable!Str10 = loctable!total
    temptable!Val10 = i
    temptable!VAL20 = 1
    temptable.Update
    loctable.MoveNext
Loop

cString = "SELECT FILE2_11.DESCA," & _
          " FILE2_11.DATE_BIRTH ," & _
          " RELATION_CODES.DESCA AS RELATION_DESCA " & _
          " FROM FILE2_11 " & _
          " INNER JOIN RELATION_CODES ON FILE2_11.RELATION = RELATION_CODES.CODE" & _
          " WHERE FILE2_11.MEMBER = " & addvalue(xCode.text)
cString = cString & " ORDER BY FILE2_11.RELATION," & _
                    " FILE2_11.DATE_BIRTH"

Set loctable = New ADODB.Recordset
Set loctable = cmd(cString, con).Execute
Do Until loctable.EOF
    temptable.AddNew
    temptable!str1 = loctable!Desca
    temptable!str2 = TurnValue(ArbString(myFormat_p(loctable!DATE_BIRTH)))
    temptable!str3 = loctable!RELATION_DESCA
    temptable!str15 = ArbString(MyOnly(mRound(nTotal)))
    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
    temptable!Val10 = i
    temptable!VAL20 = 2
    temptable.Update
    loctable.MoveNext
Loop


Set loctable = Nothing

End With
contemp.BeginTrans
contemp.CommitTrans

Report1.Reset
Report1.WindowState = crptMaximized
Report1.ReportFileName = App.Path & "\Reports\paid_install.rpt"
Report1.DataFiles(0) = tempFile

Report1.WindowShowPrintSetupBtn = True
Report1.ProgressDialog = False
Report1.CopiesToPrinter = 1

iSubreports = Report1.GetNSubreports
If (iSubreports <> 0) Then
    For i = 0 To iSubreports - 1
        sSubreportName = Report1.GetNthSubreportName(i)
        Report1.SubreportToChange = sSubreportName
        Report1.DataFiles(0) = tempFile
    Next
End If
'REPORT1.Destination = crptToPrinter
Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Function
Private Sub xYears_GotFocus()
myGotFocus xYears
End Sub
Private Sub xYears_LostFocus()
myLostFocus xYears
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Function myAddInstall(Optional nTotalMax As Double = -1) As Boolean
Dim aSql As Variant
aRet = addInstall(xCode.text, myFormat(xDate.text), con, IIf(xdoc_no.Tag = LoadMode, xdoc_no.text, ""), nTotalMax, bFawry, Val(xCard_value.text), Val(xOther.text))
If IsEmpty(retFlag(aRet, "error")) Then
    If Not IsEmpty(retFlag(aRet, "msg")) Then
        MsgBox retFlag(aRet, "msg")
    End If
    
    aSql = retFlag(aRet, "sql")
    If Not IsEmpty(aSql) Then
        prog1.Visible = True
        con.BeginTrans
        For i = 0 To UBound(aSql)
            If UBound(aSql) > 0 Then
                prog1.Value = IIf(Round(i / (UBound(aSql)), 2) > 1, 1, mRound(i / (UBound(aSql)), 2)) * 100
            End If
            On Error GoTo myerror
            con.Execute aSql(i)
        Next
        
        If bFawry Then
            Dim sForm_no As String
            sForm_no = Newflag("file6_30h", "form_no", con, "FILE6_30H.ISFAWRY = 1") & ""
            con.Execute "update file6_30h set form_no = " & sForm_no & ",form_no2 = " & (sForm_no + 1) & ",date_Cashed = [date] where doc_no = " & addvalue(retFlag(aRet, "doc_no"))
        End If
        con.CommitTrans
        
        prog1.Visible = False
        xdoc_no.text = retFlag(aRet, "doc_no")
        myUndo
        
        If bFawry Then
             MsgBox IIf(xForm_No.text <> "", " „ ”œ«œ ›Ê—Ì »‰Ã«Õ", "·„ Ì „ ”œ«œ ›Ê—Ì")
        End If
    End If
Else
    MsgBox retFlag(aRet, "error")
    Exit Function
End If
prog1.Visible = False
myAddInstall = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
