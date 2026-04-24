VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form memberPayfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ì«‰«  «⁄÷«¡ «·‰«œÌ"
   ClientHeight    =   9375
   ClientLeft      =   690
   ClientTop       =   1395
   ClientWidth     =   20250
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
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
   ScaleHeight     =   9375
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CheckBox xDoneTax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " „  «·„—«Ã⁄…"
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
      Left            =   8550
      RightToLeft     =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox xDoneTaxFilter 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "«Œ›«¡ „«  „  „—«Ã⁄ Â"
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
      Height          =   240
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1710
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Timer CardTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2385
      Top             =   2340
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1545
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   9855
      Visible         =   0   'False
      Width           =   5550
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›… «·«⁄÷«¡"
         Height          =   600
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   765
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «·’Ê—"
         Height          =   600
         Left            =   -405
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… «· Ê«»⁄"
         Height          =   600
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«÷«›… ’Ê— «· Ê«»⁄"
         Height          =   600
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   495
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command5"
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   405
         Visible         =   0   'False
         Width           =   3075
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   13050
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   3705
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   14
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Member_Pay.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   2475
         TabIndex        =   15
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Member_Pay.frx":2323
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member_Pay.frx":46EE
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   1305
         TabIndex        =   28
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Member_Pay.frx":6797
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member_Pay.frx":918C
      End
   End
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
      Height          =   1320
      Left            =   10035
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   6720
      Begin VB.TextBox xCodeInstall 
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
         Height          =   330
         Left            =   360
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "2"
         Top             =   180
         Width           =   2085
      End
      Begin VB.TextBox xNotes 
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
         Height          =   330
         Left            =   360
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   5235
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
         Height          =   330
         Left            =   360
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   5235
      End
      Begin VB.TextBox xCode 
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
         Height          =   330
         Left            =   4050
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "2"
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂÊœ «· ﬁ”Ìÿ"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·ÕÊŸ…"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂÊœ «·⁄÷Ê"
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
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈”„ «·⁄÷Ê"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   8640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   10080
      Top             =   9090
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
      Left            =   6210
      Top             =   2565
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
      Left            =   12915
      Top             =   9225
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   375
      Left            =   -1620
      Top             =   1350
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
      DragIcon        =   "Member_Pay.frx":BA25
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   10755
      TabIndex        =   6
      Tag             =   "-1"
      Top             =   8955
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin MSAdodcLib.Adodc data11 
      Height          =   420
      Left            =   10980
      Top             =   4005
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   21
      Top             =   8910
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
      Height          =   5460
      Left            =   135
      TabIndex        =   27
      Top             =   2025
      Width           =   16620
      _cx             =   29316
      _cy             =   9631
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
      ForeColorSel    =   -2147483640
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   13050
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7470
      Width           =   3660
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2745
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
         Picture         =   "Member_Pay.frx":BE67
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member_Pay.frx":E00E
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1845
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
         Picture         =   "Member_Pay.frx":10055
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member_Pay.frx":12140
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   990
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
         Picture         =   "Member_Pay.frx":1413A
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member_Pay.frx":1624B
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   90
         TabIndex        =   20
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
         Picture         =   "Member_Pay.frx":18245
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member_Pay.frx":1A469
      End
   End
End
Attribute VB_Name = "memberPayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection, aRecords As Variant
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nPhoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String, bAct As Boolean
Dim formMode As Byte
Dim oSearch As New Search, oSearchRel As New Search, oSearchClaim As New Search_empty
Dim CardTable As ADODB.Recordset
Dim TimerMode As Integer
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Dim bCheck As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdSave.Enabled = bEditRecord
cmdInform.Enabled = (nMode = LoadMode)

aRecords = retRecords(xCode.Text)
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
Sub myDefine()
xCode.Text = ""
xDesca.Text = ""
xNotes.Text = ""
xCodeInstall.Text = ""
xDoneTax.Value = 0

bCheck = False
xDoneTax.Value = 0
bCheck = True



panel1(0).Caption = ""
panel1(1).Caption = ""
panel1(2).Caption = ""
panel1(3).Caption = ""
panel1(4).Caption = ""

Fixgrd
grid1.rows = 1
myAddItem

Handlecontrols DefineMode
On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
grid1.SetFocus
Err.Clear
End Sub
Sub myProc(Optional sControl As String)
If ActiveControl.Name = cmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
    myUndo
End If
End Sub
Private Sub myload()
xCode.Text = CardTable!CODE & ""
xDesca.Text = CardTable!Desca & ""
xNotes.Text = CardTable!notes & ""
xCodeInstall.Text = CardTable!CodeInstall & ""
xDoneTax.Value = IIf(CardTable!doneTax, 1, 0)

Handlecontrols LoadMode
myloadgrd

On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
Err.Clear
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "CodeInstall", addvalue(xCodeInstall.Text))
aInsert = AddFlag(aInsert, "DoneTax", xDoneTax.Value)
con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    'aInsert = AddFlag(aInsert, "Code", addvalue(xCode.Text))
    'con.Execute addInsert(aInsert, "FILE1_10")
Else
    con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & addvalue(xCode.Text))
End If
If (Row = -1 And Row2 = -1) Or Row <> -1 Then myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdAdd_Click()
myDefine
xCode.SetFocus
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdNext_Click()
openCardtable xCode.Text, ">"
If CardTable.EOF Then openCardtable xCode.Text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardtable xCode.Text, "<"
If CardTable.EOF Then openCardtable xCode.Text, "="
myload
End Sub
Private Sub CmdFirst_Click()
openCardtable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardtable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdInform_Click()
MemberLookupAll Me, oSearch, cFilter
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
myUndo
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xCode.Tag = LoadMode Then
        grid1.SetFocus
    Else
        xCode.SetFocus
    End If
    Err.Clear
End If
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
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
    
End If
End Sub
Private Sub Form_Load()
openCon con

Set grid1.DataSource = DATA11
Fixgrd
myUndo
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.Text) Then
     If xCode.Tag = LoadMode Then
        myDefine
    Else
        xCode.Text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
        If CardTable!CODE = xCode.Text Then
            Exit Sub
        End If
    End If
    
    openCardtable xCode.Text
    If Not CardTable.EOF Then
        myload
    ElseIf xCode.Tag = LoadMode Then
        myDefine
    Else
        'xCode.Text = ""
    End If
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not ValidNum(xCode.Text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If (Not ValidNum(xCodeInstall.Text)) And Trim(xCodeInstall.Text) <> "" Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If


If Trim(xDesca.Text) = "" Then
    MsgBox "√”„ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(xCode1.Name, xcode2.Name)
CardTable.Close
Set CardTable = Nothing
End Sub
Private Sub CalcTotals()
Dim nValue
For i = 1 To grid1.rows - 2
    nValue = Val(grid1.TextMatrix(i, 2))
Next
xTotal.Caption = Format(nValue, "fixed")
End Sub
Private Function openCardtable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE1_10.*  FROM FILE1_10"
If pCode <> "" Then cWhere = "FILE1_10.CODE " & pSign & addvalue(pCode)

cFilter = ""
If sCode <> "" Then cFilter = "FILE1_10.CODE = " & addvalue(sCode)
cFilter = cFilter & turn(cFilter, " AND ") & "FILE1_10.INSTALL = 1"
If Me.xDoneTaxFilter.Value = 1 Then cFilter = cFilter & turn(cFilter, " AND ") & "FILE1_10.DoneTax = 0"
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by FILE1_10.CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by FILE1_10.CODE ASC"
End If

CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If ValidNum(pCode) Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN FILE1_10.CODE <= " & pCode & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM FILE1_10 " & turn(cFilter, " WHERE ") & cFilter

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If ValidNum(pCode) Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Text) Then
    openCardtable xCode.Text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardtable , ">"
If CardTable.EOF Then myDefine Else myload
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
Private Sub grid1_KeyPress(KeyAscii As Integer)
With grid1
If KeyAscii = 13 Then KeyAscii = 0
End With
End Sub
Private Sub Grid1_Keyup(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
With grid1
    If KeyCode = 46 And .Row <> .rows - 1 And bEditRecord Then
        If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbDefaultButton2 + vbOKCancel) Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                Dim fs As New FileSystemObject
                con.BeginTrans
                con.Execute "Delete  from file1_60 where id = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            myRemove .Row
            Grid1_EnterCell
            On Error Resume Next
            grid1.SetFocus
            Err.Clear
        End If
    ElseIf KeyCode = 13 Then
        CellPos KeyCode, .Row, .Col
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo myerror
With grid1
If Not MYVALID Then
    On Error Resume Next
    .SetFocus
    Err.Clear
    myloadgrd
    If Row < .rows - 1 Then
        .Select Row, Col
    Else
        CellPos 13, .rows - 2, .Cols - 1
    End If
    Exit Sub
End If
If Not validRow(Row) Then Exit Sub
If Row = .rows - 1 Then
    myAddItem
End If
If myreplace(Row) Then
    If xCode.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadgrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
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
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 2)) Then Exit Function
'If Not ValidDateTax(.TextMatrix(Row, 2)) Then
'    MsgBox " «—ÌŒ ”œ«œ €Ì— ’«·Õ"
'    Exit Function
'End If
If Val(.TextMatrix(Row, 3)) <= 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'With grid1
'If OldRow < 1 Then Exit Sub
'If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
'    If Not validRow(OldRow) Then
'        myRemove OldRow
'    End If
'End If
'End With
End Sub
Private Sub Grid1_EnterCell()
With grid1
If (.Col = 0 And Trim(.TextMatrix(.Row, .Cols - 1)) <> "") Then
    .Editable = flexEDNone
Else
    .Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub grid1_GotFocus()
Grid1_EnterCell
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If OldRow < 1 Then Exit Sub
If (Not validRow(.Row)) And .Row <> .rows - 1 And .TextMatrix(.Row, .Cols - 1) = "" Then
    myRemove .Row
End If
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "—ﬁ„ «·«Ì’«· €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Not IsDate(grid1.EditText) Then
        MsgBox " «—ÌŒ «·«Ì’«· €Ì— „”Ã·"
        Cancel = True
'    ElseIf Not ValidDateTax(grid1.EditText) Then
'        MsgBox " «—ÌŒ «·«Ì’«· €Ì— ’«·Õ"
'        Cancel = True
    Else
        grid1.EditText = myFormat_p(grid1.EditText)
    End If
ElseIf Col = 3 Then
    If Val(.EditText) <= 0 Then
        MsgBox "ﬁÌ„… «·«Ì’«· €Ì— „”Ã·…"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "„”·”·|" & "—ﬁ„ «·«Ì’«·|" & " «—ÌŒ «·«Ì’«·|" & "ﬁÌ„… «·«Ì’«·|" & "„·«ÕŸ« |"
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 1400
.ColWidth(3) = 1200
.ColWidth(4) = 4000
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
FixSerial
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 3 Then
    If .Col = 0 Or .Col = 1 Then
        .Col = NextEmpty(grid1, Row, Col + 1, 3)
    Else
        .Col = Col + 1
    End If
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 3)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
If grid1.rows > 2 Then
    .TextMatrix(.rows - 1, 0) = Val(grid1.TextMatrix(.rows - 2, 0)) + 1
Else
    .TextMatrix(.rows - 1, 0) = "1"
End If
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 2, Row)
        aInsert = AddFlag(Empty, "Member", addvalue(xCode.Text))
        aInsert = AddFlag(aInsert, "Form", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[Date]", addDate(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[Value]", Val(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "Notes", addstring(grid1.TextMatrix(i, 4)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_60")
        Else
            con.Execute addUpdate(aInsert, "FILE1_60", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
Dim cString As String
cString = "SELECT FILE1_60.FORM,CONVERT(VARCHAR(10),FILE1_60.DATE,111),FILE1_60.VALUE,FILE1_60.NOTES,FILE1_60.ID " & _
          " FROM FILE1_60"
cString = cString & " WHERE FILE1_60.MEMBER = " & xCode.Text
cString = cString & " ORDER BY FILE1_60.ID"
Set DATA11.Recordset = myRecordSet(cString, con)
myAddItem
Fixgrd
End With
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
FixSerial
End Sub
Private Function FoundOtheritem(grid1 As Variant, nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function

Private Sub xDoneTaxFilter_Click()
myUndo
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xCodeInstall_GotFocus()
myGotFocus xCodeInstall
End Sub
Private Sub xCodeInstall_LostFocus()
myLostFocus xCodeInstall
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub FixSerial()
Dim i As Long
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub

