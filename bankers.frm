VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form bankersfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  „ÊŸ›Ì «·»‰ﬂ"
   ClientHeight    =   7515
   ClientLeft      =   615
   ClientTop       =   1320
   ClientWidth     =   11235
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
   ScaleHeight     =   7515
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
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
      Left            =   8145
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   900
      Width           =   1320
   End
   Begin VB.TextBox xMember 
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
      Left            =   4500
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   1320
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   6480
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   35
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
         Picture         =   "bankers.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "bankers.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   36
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
         Picture         =   "bankers.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "bankers.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   37
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
         Picture         =   "bankers.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "bankers.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   38
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
         Picture         =   "bankers.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "bankers.frx":EB26
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1545
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   9855
      Visible         =   0   'False
      Width           =   5550
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›… «·«⁄÷«¡"
         Height          =   600
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   765
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «·’Ê—"
         Height          =   600
         Left            =   -405
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1485
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… «· Ê«»⁄"
         Height          =   600
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   270
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«÷«›… ’Ê— «· Ê«»⁄"
         Height          =   600
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   495
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command5"
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   405
         Visible         =   0   'False
         Width           =   3075
      End
   End
   Begin VB.Frame Frame9 
      Enabled         =   0   'False
      Height          =   2040
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2745
      Width           =   10905
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1620
         Width           =   7485
      End
      Begin VB.TextBox xMail 
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
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1260
         Width           =   7485
      End
      Begin VB.TextBox xAddress 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   7485
      End
      Begin VB.TextBox xPhone 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   7485
      End
      Begin VB.TextBox xMobil 
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
         Height          =   345
         Left            =   1845
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   7485
      End
      Begin VB.Label Label22 
         Caption         =   "„·«ÕŸ« "
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label Label18 
         Caption         =   "»—Ìœ «·Ìﬂ —Ê‰Ï"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1305
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "⁄‰Ê«‰ «·⁄÷Ê"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label Label9 
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label13 
         Caption         =   "«· ·Ì›Ê‰ «·«—÷Ì"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   1440
      End
   End
   Begin VB.Frame Frame7 
      Height          =   690
      Left            =   5805
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   0
      Width           =   5235
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "bankers.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "bankers.frx":131EE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "‰ﬁ· «·Ì «·⁄÷ÊÌ…"
         CausesValidation=   0   'False
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
         Left            =   2475
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   3960
         Picture         =   "bankers.frx":1565A
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   810
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   225
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
      Left            =   4455
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
      Left            =   2115
      Top             =   -135
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   375
      Left            =   0
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
      Caption         =   "DATA7"
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   375
      Left            =   0
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   60
      Top             =   7170
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   375
      Left            =   0
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
   Begin VB.Frame Frame6 
      Enabled         =   0   'False
      Height          =   1725
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   4770
      Width           =   10950
      Begin VB.TextBox xDate_Degree 
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
         Height          =   345
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Tag             =   "D"
         Top             =   585
         Width           =   1590
      End
      Begin VB.TextBox xDate_Job_begin 
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
         Height          =   345
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Tag             =   "D"
         Top             =   225
         Width           =   1590
      End
      Begin VB.TextBox xDate_job 
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
         Height          =   345
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Tag             =   "D"
         Top             =   945
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xJob 
         Height          =   330
         Left            =   6075
         TabIndex        =   14
         Top             =   180
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xDegree 
         Height          =   330
         Left            =   6075
         TabIndex        =   15
         Top             =   540
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xSection 
         Height          =   330
         Left            =   6075
         TabIndex        =   16
         Top             =   900
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   6075
         TabIndex        =   17
         Top             =   1260
         Width           =   3345
         _ExtentX        =   5900
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
      Begin VB.Label Label1 
         Caption         =   " «—ÌŒ «·„ƒÂ·"
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   " «—ÌŒ ‘€· «·ÊŸÌ›…"
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label Label10 
         Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "«·Õ«·…"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label24 
         Caption         =   "«·„ƒÂ·"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "«·«œ«—… «·⁄«„…"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "«·ÊŸÌ›… «·Õ«·Ì…"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   720
      Width           =   10905
      Begin VB.TextBox xId_no 
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
         Left            =   6525
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1620
         Width           =   2805
      End
      Begin VB.TextBox xDate_birth 
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
         Left            =   6525
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "D"
         Top             =   900
         Width           =   2805
      End
      Begin VB.TextBox xTitle 
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
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   2715
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
         Left            =   4365
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   4965
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   585
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xReligion 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   945
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xSocial 
         Height          =   330
         Left            =   6525
         TabIndex        =   7
         Top             =   1260
         Width           =   2805
         _ExtentX        =   4948
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
      Begin VB.Label Label8 
         Caption         =   "ﬂÊœ «·⁄÷ÊÌ…"
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "«·—ﬁ„ «·ﬁÊ„Ì"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label19 
         Caption         =   "«·Õ«·… «·«Ã „«⁄Ì…"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "«·œÌ«‰…"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "«·‰Ê⁄"
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   " «—ÌŒ «·„Ì·«œ"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label11 
         Caption         =   "«··ﬁ»"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label7 
         Caption         =   "—ﬁ„ «·ﬁÌœ"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "«·«”„"
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
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   1005
      End
   End
   Begin VB.Label xRecordNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6570
      Width           =   4020
   End
End
Attribute VB_Name = "bankersfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim fs As New FileSystemObject
Public bEdit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nphoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String
Dim formMode As Byte
Dim oSearch As New Search3, oSearchRel As New Search3
Dim CardTable As ADODB.Recordset, bAct As Boolean
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
cmdTrans.Enabled = (Trim(xMember.Text) = "")
xMember.Enabled = (Trim(xMember.Text) = "")
xCode.Enabled = bEdit = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = ""
xTitle.Text = ""
xDesca.Text = ""
xDate_birth.Text = ""
xDate_Job_begin.Text = ""
xGender.BoundText = ""
xSocial.BoundText = ""
xReligion.BoundText = ""
xId_no.Text = ""
xAddress.Text = ""
xPhone.Text = ""
xMobil.Text = ""
xMail.Text = ""
xJob.BoundText = ""
xDegree.BoundText = ""
xSection.BoundText = ""
xStatus.BoundText = ""
xDate_Job_begin.Text = ""
xDate_Degree.Text = ""
xDate_job.Text = ""

StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(4).Text = ""
Handlecontrols DefineMode
xRecordNo.Caption = "·«  ÊÃœ ”Ã·« "
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.row, 0)
    Unload oSearch
End If
myUndo
End Sub
Private Sub MyLoad()
xCode.Text = CardTable!CODE & ""
xMember.Text = CardTable!Member & ""
xTitle.Text = CardTable!Title & ""
xDesca.Text = CardTable!Desca & ""
xDate_birth.Text = myFormat_p(CardTable!DATE_BIRTH)
xGender.BoundText = CardTable!GENDER & ""
xSocial.BoundText = CardTable!Social & ""
xReligion.BoundText = CardTable!Religion & ""
xId_no.Text = CardTable!id_no & ""
xAddress.Text = CardTable!Address & ""
xPhone.Text = CardTable!Phone & ""
xMobil.Text = CardTable!MOBIL & ""
xMail.Text = CardTable!Mail & ""
xJob.BoundText = CardTable!job & ""
xDegree.BoundText = CardTable!Degree & ""
xSection.BoundText = CardTable!Section & ""
 xStatus.BoundText = CardTable!Status & ""
xDate_Job_begin.Text = myFormat_p(CardTable!Date_Job_begin)
xDate_Degree.Text = myFormat_p(CardTable!Date_Degree)
xDate_job.Text = myFormat_p(CardTable!Date_job)
Handlecontrols LoadMode

StatusBar1.Panels(1).Text = CardTable!UserName & ""
StatusBar1.Panels(2).Text = myFormat_p(CardTable!Time, True)
StatusBar1.Panels(3).Text = CardTable!UserName2 & ""
StatusBar1.Panels(4).Text = myFormat_p(CardTable!Time2, True)
xRecordNo.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
BankerLookupAll Me, oSearch, cFilter
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub

Private Sub cmdTrans_Click()
If MsgBox("‰ﬁ· «·Ì", vbDefaultButton2 + vbOKCancel) = vbOK Then
    If TransMember Then
        Inform " „ ‰ﬁ· »Ì«‰«  «·„ÊŸ› »‰Ã«Õ"
        openCardTable
        myUndo
    End If
End If
End Sub
Private Function TransMember() As Boolean
Dim aInsert As Variant

aInsert = AddFlag(aInsert, "Code", addvalue(xMember.Text))
aInsert = AddFlag(aInsert, "Title", "Title")
aInsert = AddFlag(aInsert, "Desca", "Desca")
aInsert = AddFlag(aInsert, "Gender", "Gender")
aInsert = AddFlag(aInsert, "Date_birth", "Date_Birth")
aInsert = AddFlag(aInsert, "Religion", "Religion")
aInsert = AddFlag(aInsert, "Social", "Social")
aInsert = AddFlag(aInsert, "Id_no", "Id_No")
aInsert = AddFlag(aInsert, "Address", "Address")
aInsert = AddFlag(aInsert, "Phone", "Phone")
aInsert = AddFlag(aInsert, "Mobil", "Mobil")
aInsert = AddFlag(aInsert, "Mail", "Mail")
aInsert = AddFlag(aInsert, "Job", "Job")
aInsert = AddFlag(aInsert, "Degree", "Degree")
aInsert = AddFlag(aInsert, "Section", "Section")
aInsert = AddFlag(aInsert, "Status", "Status")
aInsert = AddFlag(aInsert, "Date_Job_begin", "Date_Job_begin")
aInsert = AddFlag(aInsert, "Date_Degree", "Date_Degree")
aInsert = AddFlag(aInsert, "Date_job", "Date_job")
aInsert = AddFlag(aInsert, "Type", "Type")
aInsert = AddFlag(aInsert, "Company", "1")

End Function
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    xCode.SetFocus
    Err.Clear
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
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

Set DATA1.Recordset = myRecordSet("select * from section_Codes", con)
Set xSection.RowSource = DATA1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("select * from Job_Codes", con)
Set xJob.RowSource = DATA2
xJob.ListField = "Desca"
xJob.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("select * from Degree_Codes", con)
Set xDegree.RowSource = DATA3
xDegree.ListField = "Desca"
xDegree.BoundColumn = "Code"

Set DATA4.Recordset = myRecordSet("select * from Status_Codes", con)
Set xStatus.RowSource = DATA4
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

Set DATA5.Recordset = myRecordSet("select * from Gender_Codes", con)
Set xGender.RowSource = DATA5
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

Set DATA7.Recordset = myRecordSet("select * from religion_Codes", con)
Set xReligion.RowSource = DATA7
xReligion.ListField = "Desca"
xReligion.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (grid1.Col <> 1 And grid1.Col <> 2) Then KeyAscii = 0
End Sub
Private Sub xCode_LostFocus()
Dim MyCode As String
If Not CardTable.EOF Then MyCode = CardTable!CODE
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    MyLoad
ElseIf xCode.Tag = LoadMode Then
    If MyCode <> "" Then CardTable.Find "code = " & MyCode, , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        MyLoad
    ElseIf xCode.Tag = LoadMode Then
        mydefine
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE2_10.*,FILE1_10.CODE AS MEMBER FROM FILE2_10 LEFT JOIN FILE1_10 ON FILE2_10.CODE = FILE1_10.REGISTER"
If IsNumeric(sCode) Then cString = cString & turn(cString) & "FILE2_10.CODE = " & sCode
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY FILE2_10.code"
CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If IsNumeric(xCode.Text) Then
        CardTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    MyLoad
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xDate_Last_GotFocus()
myGotFocus xDate_Last
End Sub
Private Sub xDate_Last_LostFocus()
myLostFocus xDate_Last
myValidDate xDate_Last
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xMail_GotFocus()
myGotFocus xMail
End Sub
Private Sub xMail_LostFocus()
myLostFocus xMail
End Sub
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xPhone_GotFocus()
myGotFocus xPhone
End Sub
Private Sub xPhone_LostFocus()
myLostFocus xPhone
End Sub
Private Sub xMobil_GotFocus()
myGotFocus xMobil
End Sub
Private Sub xMobil_LostFocus()
myLostFocus xMobil
End Sub
Private Sub xId_no_GotFocus()
myGotFocus xId_no
End Sub
Private Sub xId_no_LostFocus()
myLostFocus xId_no
End Sub
Private Sub xDate_birth_GotFocus()
myGotFocus xDate_birth
End Sub
Private Sub xDate_birth_LostFocus()
myLostFocus xDate_birth
myValidDate xDate_birth
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
Private Sub xSocial_GotFocus()
myGotFocus xSocial
End Sub
Private Sub xSocial_LostFocus()
myLostFocus xSocial
If Not xSocial.MatchedWithList Then xSocial.BoundText = ""
End Sub
Private Sub xDate_Degree_GotFocus()
myGotFocus xDate_Degree
End Sub
Private Sub xDate_Degree_LostFocus()
myLostFocus xDate_Degree
myValidDate xDate_Degree
End Sub
Private Sub xDate_Job_begin_GotFocus()
myGotFocus xDate_Job_begin
End Sub
Private Sub xDate_Job_begin_LostFocus()
myLostFocus xDate_Job_begin
myValidDate xDate_Job_begin
End Sub
Private Sub xDate_job_GotFocus()
myGotFocus xDate_job
End Sub
Private Sub xDate_job_LostFocus()
myLostFocus xDate_job
myValidDate xDate_job
End Sub
Private Sub xDate_Begin_GotFocus()
myGotFocus xDate_Begin
End Sub
Private Sub xDate_Begin_LostFocus()
myLostFocus xDate_Begin
myValidDate xDate_Begin
End Sub
Private Sub xjob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xjob_LostFocus()
myLostFocus xJob
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub
Private Sub xDegree_GotFocus()
myGotFocus xDegree
End Sub
Private Sub xDegree_LostFocus()
myLostFocus xDegree
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub xStatus_GotFocus()
myGotFocus xStatus
End Sub
Private Sub xStatus_LostFocus()
myLostFocus xStatus
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub
