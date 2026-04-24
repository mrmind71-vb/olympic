VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form meetMemberfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄÷«¡ «·Ã„⁄Ì… «·⁄„Ê„Ì…"
   ClientHeight    =   10980
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   18060
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
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10980
   ScaleWidth      =   18060
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fmDirect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   510
      Left            =   12780
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   3795
      TabIndex        =   11
      Top             =   1935
      Width           =   3795
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2880
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   915
         _ExtentX        =   1614
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
         Picture         =   "meet_member.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "meet_member.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1890
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   45
         Width           =   960
         _ExtentX        =   1693
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
         Picture         =   "meet_member.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "meet_member.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   990
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
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
         Picture         =   "meet_member.frx":82D3
         Caption         =   " «·Ì"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "meet_member.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   915
         _ExtentX        =   1614
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
         Picture         =   "meet_member.frx":C3DE
         Caption         =   "«ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "meet_member.frx":E602
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   1725
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   5100
      Begin VB.Label xCode_Zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1260
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·Ã„⁄Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ê”„ «·Ã„⁄Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label xSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂÊœ «·Ã„⁄Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "»Ì«‰ «·Ã„⁄Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   1155
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   13950
      Top             =   7065
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   225
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   405
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1217
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
      Picture         =   "meet_member.frx":106D3
      Caption         =   "Œ—ÊÃ"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   1890
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   405
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1217
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
      Picture         =   "meet_member.frx":129F6
      Caption         =   "ÿ»«⁄… «·«⁄÷«¡"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "meet_member.frx":14D6C
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   45
      Width           =   7530
      Begin VB.TextBox xMember2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Tag             =   "igkeypress"
         Top             =   225
         Width           =   1275
      End
      Begin VB.TextBox xDesca_Member 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "igkeypress"
         Top             =   585
         Width           =   6315
      End
      Begin VB.TextBox xMember 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5130
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "igkeypress"
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Õ Ì"
         Height          =   240
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·«”„"
         Height          =   240
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   945
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   2670
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8250
      Left            =   225
      TabIndex        =   23
      Top             =   1125
      Width           =   12525
      _cx             =   22093
      _cy             =   14552
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   27
      Top             =   10515
      Width           =   18060
      _ExtentX        =   31856
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         TabIndex        =   32
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
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   12825
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2835
      Width           =   5190
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Õ«÷—Ì‰"
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
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "·„ ÌÕ÷—Ê«"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„⁄ –—Ì‰"
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
         Index           =   3
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   225
         Width           =   1005
      End
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   690
      Left            =   3600
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   405
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1217
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
      Picture         =   "meet_member.frx":16EEF
      Caption         =   "»ÕÀ"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "meet_member.frx":192BA
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   39
      Top             =   10365
      Visible         =   0   'False
      Width           =   18060
      _ExtentX        =   31856
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3915
      Width           =   5235
      Begin VB.CommandButton cmdExel 
         Height          =   600
         Left            =   45
         Picture         =   "meet_member.frx":1B2B4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   2625
      End
      Begin Threed.SSCommand cmdPdf 
         Cancel          =   -1  'True
         Height          =   600
         Left            =   2610
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1058
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
         Picture         =   "meet_member.frx":1DA9F
         Caption         =   "Pdf ÿ»«⁄…"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "meet_member.frx":2006A
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   540
      Top             =   -90
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid grdExcel 
      Height          =   8250
      Left            =   225
      TabIndex        =   44
      Top             =   1125
      Visible         =   0   'False
      Width           =   12525
      _cx             =   22093
      _cy             =   14552
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Label xCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3555
      Width           =   5235
   End
   Begin VB.Label xRecord_No 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   12825
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2475
      Width           =   3750
   End
End
Attribute VB_Name = "meetMemberfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection, bCheck As Boolean, oSearchYear As New Search_empty
Dim cList As String, aHeader()
Dim formMode As Byte
Dim oSearch As New Search
Dim CardTable As ADODB.Recordset

Sub Handlecontrols(nMode)
bEditRecord = bedit
aRecords = retRecords(xcode_zero.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

xRecord_No.Caption = ArbString("”Ã· " & nRecord & " „‰ " & nRecords)

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sCode = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sCode = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sCode = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sCode = ""
End Sub
Sub mydefine()
xCode.Caption = ""
xcode_zero.Caption = ""
xdesca.Caption = ""
xDate.Caption = ""
xSeason.Caption = ""
'xClosed.Value = 0
Handlecontrols DefineMode
End Sub
Sub myProc()
xMember.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xCodeDesca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
myLoadGrd
End Sub
Private Sub myload()
xCode.Caption = CardTable!code & ""
xcode_zero.Caption = CardTable!CODE_ZERO & ""
xdesca.Caption = CardTable!Desca & ""
xDate.Caption = myFormat_p(CardTable!Date)
'cmdYear.Tag = CardTable!SEASON & ""
xSeason.Caption = GetField("SELECT DESCA FROM YEARS_CODES WHERE CODE = " & addvalue(CardTable!SEASON & ""), con)
bCheck = True
'xClosed.Value = IIf(CardTable!Closed, 1, 0)
bCheck = False
myLoadGrd
Handlecontrols LoadMode
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "—ﬁ„ «·⁄÷ÊÌ…|" & "«·«”„|" & "«·’›…|" & "«· „«„|" & "«·ﬂÊœ|" & "Id"
.ColWidth(0) = 1200
.ColWidth(1) = 4000
.ColWidth(2) = 2000
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.MergeCol(0) = True
.MergeCells = flexMergeFree
.ColComboList(3) = cList
.ColHidden(.Cols - 1) = True
.ColHidden(.Cols - 2) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub cmdAdd_Click()
mydefine
'xDesca.SetFocus
End Sub

Private Sub cmdAddMember_Click()
Me.MousePointer = 11
If AddMember Then
    Me.MousePointer = 0
    Inform "‰„ «÷«›… «⁄÷«¡«·Ã„⁄Ì… »‰Ã«Õ"
Else
    Me.MousePointer = 0
    Inform "·„   „ «÷«›… «⁄÷«¡«·Ã„⁄Ì… »‰Ã«Õ"
End If
End Sub

Private Sub CmdDel_Click()
Dim nDelete As Long
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From MEETING Where code = " & xCode.Caption, nDelete
    con.CommitTrans
    openCardTable xCode.Caption, "<="
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
Private Sub cmdExel_Click()
Dim cString As String
cString = doPrint(False, True)

If cString = "" Then Exit Sub

Set grdExcel.DataSource = data10
Set data10.Recordset = mySet(cString, con)

With grdExcel
.ColWidth(0) = 800
.ColWidth(1) = 2000
.ColWidth(2) = 1000
.ColWidth(3) = 3000
.ColWidth(4) = 1700
.ColWidth(5) = 1700
.ColWidth(6) = 1500

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«”„ «·‰«œÌ"
.TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
.TextMatrix(0, 3) = "«·«”„ «·—»«⁄Ì"
.TextMatrix(0, 4) = "«·—ﬁ„ «·ﬁÊ„Ì"
.TextMatrix(0, 5) = "«·ÊŸÌ›…"
.TextMatrix(0, 6) = "«· ·Ì›Ê‰ «·„Õ„Ê·"
.ColHidden(2) = True
End With

ToFileExel2 grdExcel, , , , , 1, , , , , , Me, Array(xdesca.Caption)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdInform_Click()
'MeetingLookup Me, oSearch
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
myUndo
End Sub
Private Sub CmdNext_Click()
openCardTable xcode_zero.Caption, ">"
If CardTable.EOF Then openCardTable xCode.Caption
myload
End Sub
Private Sub cmdPdf_Click()
doPrint True
End Sub
Private Function doPrint(Optional pPdf As Boolean = False, Optional bExcel As Boolean = False) As String
Dim aHeader(12)
Dim cString As String
Dim cStringWhere As String
cStringWhere = "SELECT MEM_MEETING.CODE FROM  MEM_MEETING  WHERE MEM_MEETING.CODE_MEETING = " & addvalue(xCode.Caption)
If Option1(1).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 0"
    aHeader(0) = Option1(1).Caption
ElseIf Option1(2).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 1"
    aHeader(0) = Option1(2).Caption
ElseIf Option1(3).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 2"
    aHeader(0) = Option1(3).Caption
End If
cString = "Select ROW_NUMBER() OVER(ORDER BY FILE1_10.CODE ASC) AS Row_Number ,'‰«œÌ «·«Ê·Ì„»Ì «·”ﬂ‰œ—Ì' as club_desca,file1_10.CODE,FILE1_10.DESCA,FILE1_10.ID_NO,JOB_CODES.DESCA AS Job_Desca,FILE1_10.MOBIL" & _
          " From File1_10  LEFT JOIN JOB_CODES ON FILE1_10.JOB = JOB_CODES.CODE" & _
          " WHERE FILE1_10.CODE  IN (" & cStringWhere & ")"


If cWhere <> "" Then cStringWhere = cStringWhere & " AND " & cWhere

If bExcel Then
    doPrint = cString & " ORDER BY FILE1_10.CODE"
    Exit Function
End If



Dim temptable As New ADODB.Recordset, sourcetable As New ADODB.Recordset

Me.MousePointer = 11
Set sourcetable = myCmd(cString, con)

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!val1 = sourcetable!code
    temptable!str1 = sourcetable!code
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(sourcetable!ID_NO)
    temptable!str4 = TurnValue(sourcetable!JOB_desca)
    temptable!str5 = sourcetable!Mobil
    temptable!str7 = sourcetable!club_desca
        
    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(xdesca.Caption)
    temptable!str12 = TurnValue(retHeader(aHeader, 0, 1))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    If pPdf Then
        FixPrinter Report1, 1
        Report1.Destination = crptToPrinter
    End If
    Report1.ReportFileName = sPath_App & "\REPORTS\REPORT26.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
    Me.MousePointer = 0
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function

Private Sub CmdPrevious_Click()
openCardTable xcode_zero.Caption, "<"
If CardTable.EOF Then openCardTable xCode.Caption
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
Years_LookupAll Me, oSearchYear, , cmdYear.Tag <> ""
End Sub

Private Sub cmdPrint_Click()
Dim nRate As Double
Dim nwidth As Double
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then
        nwidth = grid1.ColWidth(i) + nwidth
    End If
Next
nRate = 11500 / nwidth
Set PrintGrdNew.myForm = Me
PrintGrdNew.doPrint grid1, nRate, 0, Me.Caption & " " & xdesca.Caption & turn(aHeader(0), " - ") & aHeader(0), , , , False, False, 11, , aRow, Array(0)
PrintGrdNew.Show 1
End Sub

Private Sub cmdSearch_Click()
myLoadGrd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        If ActiveControl.Tag <> "igkeypress" Then
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If ActiveControl.Tag <> "igkeypress" Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()

openCon con
bedit = True

Set grid1.DataSource = data1
cList = StrList2("Select code,desca from meeting_codes order by code")
myUndo
End Sub
Private Sub xCode_LostFocus()
'myLostFocus xCode
'If Not ValidNum(xcode.caption) Then
'     If xCode.Tag = LoadMode Then
'        mydefine
'    Else
'        xcode.caption = ""
'    End If
'Else
'    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
'        If CardTable!CODE = xcode.caption Then
'            Exit Sub
'        End If
'    End If
'
'    openCardTable xcode.caption
'    If Not CardTable.EOF Then
'        myload
'    ElseIf xCode.Tag = LoadMode Then
'        mydefine
'    Else
'        'xcode.caption = ""
'    End If
'End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
'If Not ValidInt(xcode.caption) Then
'    If Not igMsg Then MsgBox "«·ﬂÊœ €Ì— „”Ã·", , systemName
'    Exit Function
'End If
'
'If Not IsDate(xDate.Text) Then
'    If Not igMsg Then MsgBox " «—ÌŒ «·Ã„⁄Ì… €Ì— „”Ã·…", , systemName
'    Exit Function
'End If
'
'If Not IsDate(xdate2.Text) Then
'    If Not igMsg Then MsgBox " «—ÌŒ «‰ Â«¡ «·”œ«œ €Ì— „”Ã·", , systemName
'    Exit Function
'End If
'
'If Not ValidNum(cmdYear.Tag) Then
'    MsgBox "„Ê”„ «·”œ«œ €Ì— „”Ã·", , systemName
'    Exit Function
'End If
'
'If xDesca.Text = "" Then
'    If Not igMsg Then MsgBox "»Ì«‰ «·Ã„⁄Ì… €Ì— „”Ã·", , systemName
'    Exit Function
'End If
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
If ValidNum(xCode.Caption) Then
    openCardTable xCode.Caption
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
On Error Resume Next
xMember.SetFocus
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
Private Function retMeetCode(pMeet As String, pMember As String, pRel As String) As String
retMeetCode = RetZero(pMeet, 2) & "-" & RetZero(pMember, 10) & "-" & RetZero(pRel, 2)
End Function
Private Sub myLoadGrd()
ReDim aHeader(5)
With grid1
Dim cString As String
cString = "SELECT MEM_MEETING.CODE,MEMBER_WIFE.DESCA, MEMBER_WIFE.RELATION_DESCA, MEM_MEETING.TYPE,MEM_MEETING.CODE_REL, MEM_MEETING.ID" & _
          " FROM MEM_MEETING INNER JOIN MEMBER_WIFE ON MEM_MEETING.CODE = MEMBER_WIFE.CODE AND MEM_MEETING.CODE_REL = MEMBER_WIFE.CODE_REL"
cString = cString & " WHERE MEM_MEETING.CODE_MEETING = " & addvalue(xCode.Caption)
If ValidNum(xMember.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.CODE " & IIf(ValidNum(xMember2.text), " >= ", " = ") & addvalue(xMember.text)
End If

If ValidNum(xMember2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.CODE  <= " & addvalue(xMember2.text)
End If

If Option1(1).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 0"
    aHeader(0) = Option1(1).Caption
ElseIf Option1(2).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 1"
    aHeader(0) = Option1(2).Caption
ElseIf Option1(3).Value Then
    cWhere = cWhere & turn(cWhere, " AND ") & " MEM_MEETING.TYPE = 2"
    aHeader(0) = Option1(3).Caption
End If

If Trim(xDesca_Member.text) <> "" Then
    cWhere = cWhere & turn(cWhere, " AND ") & MyParnAnd(xDesca_Member.text, "MEMBER_WIFE.DESCA ")
End If
If cWhere <> "" Then cString = cString & " AND " & cWhere
cString = cString & " ORDER BY MEM_MEETING.CODE,MEM_MEETING.CODE_REL"
Set data1.Recordset = myRecordSet(cString, con)
Fixgrd
xCount.Caption = "⁄œœ «·«⁄÷«¡ : " & grid1.rows - 1
End With
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
myreplace Row
End Sub
Private Function myreplace(Optional Row As Long = -1) As Boolean
con.BeginTrans
On Error GoTo myerror
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Sub grid1_DblClick()
If grid1.Row < 1 Then Exit Sub
If grid1.Col <> 3 Then
    nValue = mRound(grid1.TextMatrix(grid1.Row, 3)) + 1
    nValue = IIf(nValue > 2, 0, nValue)
    grid1.TextMatrix(grid1.Row, 3) = nValue
    myreplaceGrd grid1.Row
    If grid1.Col = 0 And grid1.Row < grid1.rows - 1 Then
        If grid1.TextMatrix(grid1.Row, 0) = grid1.TextMatrix(grid1.Row + 1, 0) Then
            grid1.TextMatrix(grid1.Row + 1, 3) = nValue
            myreplaceGrd grid1.Row + 1
        End If
    End If
End If
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 3 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Trim(grid1.EditText) = "" Then
    Cancel = True
End If
End Sub

Private Sub Option1_Click(Index As Integer)
myLoadGrd
End Sub

Private Sub xDesca_Member_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    myLoadGrd
End If
End Sub

Private Sub xMember_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
ElseIf KeyCode = 13 Then
    myLoadGrd
End If
End Sub

Private Sub xMember_LostFocus()
xCodeDesca.Caption = ""
If ValidNum(xMember.text) Then
    aMember = Member_Load(xMember.text, , con)
    If Not IsEmpty(aMember) Then
        xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
    Else
        xCodeDesca.Caption = ""
    End If
Else
    xMember.text = ""
End If
End Sub

Private Sub xMember2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    myLoadGrd
End If
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 2, Row)
        aInsert = AddFlag(aInsert, "TYPE", mRound(grid1.TextMatrix(i, 3)))
        con.Execute addUpdate(aInsert, "MEM_MEETING", "ID = " & addstring(grid1.TextMatrix(i, .Cols - 1)))
    Next
End With
myreplaceGrd = True
End Function

