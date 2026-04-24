VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reportfrm19 
   Caption         =   "»Ì«‰«  √⁄÷«¡ «·Ã„⁄Ì… «·⁄„Ê„Ì… ··‰«œÏ"
   ClientHeight    =   5880
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   9210
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9210
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   35
      Top             =   5685
      Visible         =   0   'False
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4470
      Left            =   585
      TabIndex        =   34
      Top             =   5850
      Visible         =   0   'False
      Width           =   13560
      _cx             =   23918
      _cy             =   7885
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      Rows            =   1
      Cols            =   3
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   495
      Visible         =   0   'False
      Width           =   5010
      Begin VB.CheckBox chkOtherSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "√⁄÷«¡ „”œœÌ‰ ›Ï €Ì— ”‰… «·Ã„⁄Ì…"
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
         Height          =   270
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   3345
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1545
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2745
      Width           =   8925
      Begin VB.TextBox xHeader 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   270
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   495
         Width           =   6855
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   4230
         TabIndex        =   24
         Top             =   135
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã „«⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄‰Ê«‰ «· ﬁ—Ì—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "”œ«œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1305
      Width           =   4020
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   1590
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1590
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   330
         Index           =   0
         Left            =   585
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "”œœ „Ê”„"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ì"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ·«· «·› —… „‰"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1680
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1035
      Width           =   5055
      Begin VB.CommandButton cmdAge 
         Caption         =   "«“Ê«Ã ·„ Ì’·Ê« ··”‰ «·„Õœœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1125
         Width           =   4920
      End
      Begin VB.CommandButton cmdPrintNoPay 
         Caption         =   "«⁄÷«¡ €Ì— „”œœÌ‰ «·„Ê”„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   135
         Width           =   2535
      End
      Begin VB.CommandButton cmdDied 
         Caption         =   "⁄÷Ê „ Ê›Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   2355
      End
      Begin VB.CommandButton cmdPrintYear 
         Caption         =   "√⁄÷«¡ ·„ Ì„— ⁄·ÌÂ„ ”‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   2355
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«⁄÷«¡ Õ«›ŸÌ «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4275
      Width           =   8070
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "REPORT19.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "⁄—÷"
         Top             =   720
         Width           =   7980
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "REPORT19.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   135
         Width           =   1635
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   555
         Left            =   5850
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   979
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
         Picture         =   "REPORT19.frx":4C57
         Caption         =   "ÿ»«⁄…  ﬁ—Ì— «·Ã„⁄Ì…"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "REPORT19.frx":7C55
      End
      Begin Threed.SSCommand cmdPdf 
         Cancel          =   -1  'True
         Height          =   555
         Left            =   1710
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
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
         Picture         =   "REPORT19.frx":9DD8
         Caption         =   "Pdf ÿ»«⁄…"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "REPORT19.frx":C3A3
      End
      Begin Threed.SSCommand cmdPrint_Test 
         Height          =   555
         Left            =   3555
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   979
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
         Picture         =   "REPORT19.frx":E526
         Caption         =   "ÿ»«⁄…  ﬁ—Ì— «·„—«Ã⁄…"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "REPORT19.frx":10B2A
      End
      Begin Threed.SSCommand cmdTxt 
         Height          =   555
         Left            =   7110
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   855
         Visible         =   0   'False
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   979
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "⁄„· „·› ‰’Ì"
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "REPORT19.frx":12CAD
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   4020
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox xCode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1500
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   585
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·Ã„⁄Ì… «·⁄„Ê„Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ ⁄÷ÊÌ… „‰ —ﬁ„"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï —ﬁ„"
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
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   540
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1950
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   450
      Top             =   360
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   0
      Top             =   0
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
End
Attribute VB_Name = "reportfrm19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty
Dim aHeader()
Private Sub cmdAge_Click()
DoPrintAge
End Sub
Private Sub cmdDied_Click()
DoPrintDied
End Sub
Private Sub cmdExel_Click()
Dim cString As String
If Not MYVALID Then Exit Sub

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

Set grid1.DataSource = data10
Set data10.Recordset = myCmd("[dbo].[sp_meet_member_phone]", con, adStoredProc, aPrm, 300)


With grid1
.ColWidth(0) = 2000
.ColWidth(1) = 5000
.ColWidth(1) = 5000
.TextMatrix(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
.TextMatrix(0, 1) = "«·«”„"
.TextMatrix(0, 2) = "—ﬁ„ «·„Ê»«Ì·"
.TextMatrix(0, 3) = "‰Ê⁄ «·⁄÷ÊÌ…"
.ColHidden(4) = True
End With

Dim sHeader As String, nMargin As Integer
sHeader = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.text)

Dim aSplit As Variant
aSplit = AddFlag(aSplit, "title_col", "A:B")
aSplit = AddFlag(aSplit, "title_row", "1:1")
aSplit = AddFlag(aSplit, "center_header", sHeader)
ToFileExel grid1, , , , , , , , aSplit, , , Me, nMargin
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_pdf_Click()
'Report03_1 , True
End Sub

Private Sub CmdOk_Click()
doprintNew
End Sub

Private Sub cmdPdf_Click()
doprintNew , True
End Sub

Private Sub cmdPrint_test_Click()
doprintNew True
End Sub

Private Sub cmdPrintNoPay_Click()
DoPrintNoPaid
End Sub


Private Sub cmdPrintYear_Click()
DoPrintYear
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
End Sub

Private Sub cmdTxt_Click()
createText 1000
End Sub

Private Sub cmdYear_Click(Index As Integer)
Years_LookupAll Me, oSearchYear, , cmdYear(Index).Tag <> ""
End Sub
Sub myProc()
If ActiveControl.Name = cmdYear(0).Name Then
    ActiveControl.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    ActiveControl.Caption = IIf(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0) = "", "«Œ «— «·„Ê”„", oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    oSearchYear.Hide
End If
End Sub
Private Sub Command4_Click()
doPrintSafe
End Sub

Private Sub Command7_Click()
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("select '«Ã „«⁄ «Ê·' AS DESCA UNION ALL SELECT '«Ã „«⁄ À«‰Ì'", con)
Set xType.RowSource = data1
xType.ListField = "Desca"
xType.BoundColumn = "Desca"
LoadText Me
If cmdYear(0).Tag <> "" Then cmdYear(0).Caption = GetField("Select desca from years_codes where code = " & addvalue(cmdYear(0).Tag), con) & ""
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If

If cmdYear(0).Tag = "" Then
    MsgBox "«·„Ê”„ €Ì— „Õœœ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doPrint(Optional bTest As Boolean = False, Optional pPdf As Boolean = False)
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub


cString = "SELECT FILE1_10.CODE,FILE1_10.DESCA,FILE6_20H.DATE,FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,FILE1_10.DATE_BEGIN,FILE1_10.DATE_BIRTH,FILE6_20H.TOTAL" & _
           " FROM File1_10 INNER JOIN FILE6_20H ON FILE6_20H.DOC_NO = dbo.f_meeting_doc(FILE1_10.CODE," & cmdYear(0).Tag & ")" & _
           " WHERE FILE1_10.DIED = 0"

cWhere = "FILE1_10.Date_Begin <= " & DateSq(DateAdd("yyyy", -1, xDate.text))
cWhere = cWhere & " and dbo.fn_is_safe(file1_10.code) = 0"

If chkOtherSeason.Value = 1 Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.YEAR_CODE <> " & cmdYear(0).Tag
End If

If ValidNum(xCode1.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code " & IIf(ValidNum(xCode2.text), " >= ", " = ") & addvalue(xCode1.text)
End If

If ValidNum(xCode2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE1_10.code <= " & addvalue(xCode1.text)
End If

If IsDate(xdate1.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.DATE >= " & DateSq(xdate1.text)
End If

If IsDate(xdate2.text) Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.DATE <= " & DateSq(xdate2.text)
End If


If cWhere <> "" Then cString = cString & " AND " & cWhere

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

Me.MousePointer = 11

Set sourcetable = New ADODB.Recordset
'sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
Set sourcetable = mySet(cString, con)
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    If Trim(xHeader.text) = "" Then
        temptable!str11 = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.text)
    Else
        temptable!str11 = TurnValue(xHeader.text)
    End If
    temptable!str12 = TurnValue(xType.text)
    temptable!str10 = !doc_no
    temptable!val1 = !code
    temptable!val2 = 0
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(!date_begin)))
    If bTest Then
        temptable!str4 = TurnValue(ArbString(myFormat_p(!Date)))
        temptable!str5 = !FORM_NO
    End If
    temptable!val4 = !total
    temptable.Update
    sourcetable.MoveNext
Loop
End With

cString = "SELECT FILE1_11.Member,file1_11.code,File1_11.Desca,FILE1_11.DATE_BEGIN,FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,FILE6_20H.DATE,FILE6_20H.TOTAL" & _
          " from file1_11 inner join file1_10 on file1_11.member = file1_10.code INNER JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE" & _
           " AND FILE6_20H.doc_no = dbo.f_meeting_doc(FILE1_10.CODE," & cmdYear(0).Tag & ")"
cWhere = cWhere & turn(cWhere, " AND ") & "FILE1_11.RELATION  = 1"
cWhere = cWhere & " AND " & "FILE1_11.Date_Begin <= " & DateSq(DateAdd("yyyy", -1, xDate.text))
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Set sourcetable = Nothing
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.text)
    temptable!str12 = TurnValue(xType.text)
    temptable!str2 = sourcetable!Desca
    temptable!val1 = sourcetable!MEMBER
    temptable!str10 = !doc_no
    temptable!str3 = TurnValue(ArbString(myFormat_p(!date_begin)))
    If bTest Then
        temptable!str4 = TurnValue(ArbString(myFormat_p(!Date)))
        temptable!str5 = !FORM_NO
    End If
    
    temptable!val2 = 1
    temptable!val4 = !total
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery

On Error GoTo myerror
If temptable.BOF And temptable.EOF Then
    Me.MousePointer = 0
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    If bTest Then
        Report1.ReportFileName = sPath_App & "\REPORTS\report19_t.rpt"
    ElseIf pPdf Then
        FixPrinter Report1, 1
        Report1.ReportFileName = sPath_App & "\REPORTS\report19_pdf.rpt"
        Report1.Destination = crptToPrinter
    Else
        Report1.ReportFileName = sPath_App & "\REPORTS\report19.rpt"
    End If
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
    Me.MousePointer = 0
End If

Set temptable = Nothing
Set sourcetable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub doprintNew(Optional bTest As Boolean = False, Optional pPdf As Boolean = False)
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim i As Long
Dim cString As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member]", con, adStoredProc, aPrm, 300)

nRecordcount = sourcetable.RecordCount
prog1.Visible = True
With sourcetable
Do Until sourcetable.EOF
    i = i + 1
    prog1.Value = IIf(Round(i / (nRecordcount), 2) > 1, 1, Round(i / (nRecordcount), 2)) * 100
    temptable.AddNew
    If Trim(xHeader.text) = "" Then
        temptable!str11 = "«·√⁄÷«¡ «·–Ì‰ ·Â„ Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.text)
    Else
        temptable!str11 = TurnValue(xHeader.text)
    End If
    temptable!str12 = TurnValue(xType.text)
    temptable!str10 = !doc_no
    temptable!val1 = !code
    
    temptable!val2 = !relation
    
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(!date_begin)))
    temptable!str4 = TurnValue(ArbString(myFormat_p(!Date)))
    temptable!str5 = !FORM_NO
    temptable!val4 = !total
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery

Me.MousePointer = 0
prog1.Visible = False

On Error GoTo myerror
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    If bTest Then
        Report1.ReportFileName = sPath_App & "\REPORTS\report19_t.rpt"
    ElseIf pPdf Then
        FixPrinter Report1, 1
        Report1.ReportFileName = sPath_App & "\REPORTS\report19_pdf.rpt"
        Report1.Destination = crptToPrinter
    Else
        Report1.ReportFileName = sPath_App & "\REPORTS\report19.rpt"
    End If
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub DoPrintYear()
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub
    
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member_no_year]", con, adStoredProc, aPrm, 300)

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    If Trim(xHeader.text) = "" Then
        temptable!str11 = "«·√⁄÷«¡ ·„ Ì„— ⁄·Ì ⁄÷ÊÌ Â„ ”‰… Ê·Ì” Õﬁ Õ÷Ê— «·Ã„⁄Ì… «·⁄„Ê„Ì… «·„‰⁄ﬁœ… ›Ì " & myFormat_p(xDate.text)
    Else
        temptable!str11 = TurnValue(xHeader.text)
    End If
    temptable!str12 = TurnValue(xType.text)
    temptable!str10 = !doc_no
    temptable!val1 = !code
    temptable!val2 = !relation
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(!date_begin)))
    temptable!str4 = IIf(!relation = 0, "⁄÷Ê", "«“Ê«Ã")
    temptable!str5 = !FORM_NO
    temptable!val4 = !total
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    Report1.ReportFileName = sPath_App & "\REPORTS\report19_1.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub DoPrintAge()
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub
    
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member_no_age]", con, adStoredProc, aPrm, 300)

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = "«“Ê«Ã ·„ Ì „Ê« 21 ”‰… „‰–  «—ÌŒ «·Ã„⁄Ì… «·⁄„Ê„Ì… : " & myFormat_p(xDate.text)
    temptable!str2 = sourcetable!Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_BIRTH)))
'    temptable!Str4 = TurnValue(ArbString(myFormat_p(sourceTable!Date)))
    temptable!val1 = sourcetable!code
    temptable!val2 = 0
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery
Me.MousePointer = 0
    
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    
    Report1.ReportFileName = sPath_App & "\REPORTS\report19_5.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub DoPrintNoPaid()
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String, cWhereDate As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext
    
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 0)

Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member_unpaid]", con, adStoredProc, aPrm, 300)

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = "««·«⁄÷«¡ «·–Ì‰ ·„ Ì”œœÊ« Œ·«· «·› —… " & BetweenString(myFormat_p(xdate1.text), myFormat_p(xdate2.text))
    temptable!val1 = !code
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!Date)))
    temptable!str4 = "⁄÷Ê"
    temptable!val2 = 0
    temptable.Update
    sourcetable.MoveNext
Loop
End With

Me.MousePointer = 11

If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    Report1.ReportFileName = sPath_App & "\REPORTS\report19_2.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub DoPrintDied()
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String, cWhereDate As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub

contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext
    
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 1)
aPrm = AddFlag(aPrm, "SAFE", 0)


Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member_died]", con, adStoredProc, aPrm, 300)
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = ArbString("«·Ã„⁄Ì… «·⁄„Ê„Ì… («⁄÷«¡ „ Ê›ÌÌ‰) ") & BetweenString(myFormat_p(xdate1.text), myFormat_p(xdate2.text))
    temptable!val1 = !code
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!Date)))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Me.MousePointer = 0

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    Report1.ReportFileName = sPath_App & "\REPORTS\report19_3.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub doPrintSafe()
Dim temptable As New ADODB.Recordset, sourcetable As ADODB.Recordset
Dim cString As String, cWhereDate As String
ReDim aHeader(5)
If Not MYVALID Then Exit Sub
    
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext
    
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "SEASON", cmdYear(0).Tag)
aPrm = AddFlag(aPrm, "DATE", myFormat_sp(xDate.text))
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xdate2.text))
aPrm = AddFlag(aPrm, "CODE1", TurnValue(xCode1.text))
aPrm = AddFlag(aPrm, "CODE2", TurnValue(xCode2.text))
aPrm = AddFlag(aPrm, "DIED", 0)
aPrm = AddFlag(aPrm, "SAFE", 1)

Me.MousePointer = 11

Set sourcetable = myCmd("[dbo].[sp_meet_member_save]", con, adStoredProc, aPrm, 300)

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = "«·«⁄÷«¡ «·Õ«›ŸÌ «·⁄÷ÊÌ… Ê ‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ «·Ã„⁄Ì…"
    temptable!val1 = !code
    temptable!val2 = 0
    temptable!str2 = !Desca
    temptable!str3 = TurnValue(ArbString(myFormat_p(sourcetable!Date)))
    temptable!str4 = TurnValue(ArbString(myFormat_p(sourcetable!Date)))
    temptable.Update
    sourcetable.MoveNext
Loop
End With

Me.MousePointer = 0

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.Reset
    Report1.ProgressDialog = False
    Report1.WindowState = crptMaximized
    Report1.ReportFileName = sPath_App & "\REPORTS\report19_4.rpt"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me, , Array(xDate.Name, xdate1.Name, xdate2.Name, cmdYear(0).Name, xType.Name, xHeader.Name, xCode1.Name, xCode2.Name)
On Error Resume Next
con.Close
Set con = Nothing
Err.Clear
End Sub

Private Sub xHeader_GotFocus()
myGotFocus xHeader
End Sub
Private Sub xHeader_LostFocus()
myLostFocus xHeader
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub createText2(nRecordsFile)
'Dim TextLine As String, cData As String, aInsert As Variant, cFile As String, cName As String
'Dim cString As String
'
'cString = getTableString
'If cString = "" Then Exit Sub
'Dim loctable As New ADODB.Recordset
'Set loctable = mySet(cString, con)
'nRecords = loctable.RecordCount
'If loctable.EOF And loctable.EOF Then Exit Sub
'
''nFiels = nRecords
''nRecords = IIf(Int(nRecordsFiles / nRecords) < mRound(nRecordsFiles / nRecords, 2), Int(nRecordsFiles / nRecords) + 1, Int(nRecordsFiles / nRecords))
''
''Do Until locTable.EOF
''
''    locTable.MoveNext
''Loop
'Dim aString As Variant
'Do Until loctable.EOF
'    nRecord = nRecord + 1
'    cInsert = cInsert & turn(cInsert, vbCrLf) & loctable!CODE & ";2" & loctable!Mobil & ";" & loctable!desca
'    If nRecord = nRecordsFile Then
'        aString = AddFlag(aString, cInsert)
'        cInsert = ""
'        nRecord = 0
'    End If
'    loctable.MoveNext
'Loop
'
'If cInsert <> "" Then
'    aString = AddFlag(aString, cInsert)
'End If
'
'Dim fs As New FileSystemObject
'
'For i = 0 To UBound(aString)
'    cName = (i + 1)
'    cFile = App.Path & "\txt\" & cName & ".txt"
'
'    On Error GoTo myerror
'    Open cFile For Output As #1   ' Open file.
'    Print #1, aString(i)        ' Read line into variable.
'    MsgBox "DONE " & "(" & (i + 1) & ")"
'    Close #1   ' Close file.
'Next
'Exit Sub
'myerror:
'MsgBox Err.Description
'Err.Clear
End Sub
Private Sub createText(nRecordsFile)
Dim TextLine As String, cData As String, aInsert As Variant, cFile As String, cName As String
Dim cString As String

cString = getTableString
If cString = "" Then Exit Sub
Dim loctable As New ADODB.Recordset
Set loctable = mySet(cString, con)
nRecords = loctable.RecordCount
If loctable.EOF And loctable.EOF Then Exit Sub

Dim aString As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    cInsert = cInsert & turn(cInsert, vbCrLf) & loctable!code & ";2" & loctable!Mobil & ";" & loctable!Desca
    If nRecord = nRecordsFile Then
        aString = AddFlag(aString, cInsert)
        cInsert = ""
        nRecord = 0
    End If
    loctable.MoveNext
Loop

If cInsert <> "" Then
    aString = AddFlag(aString, cInsert)
End If

Dim fs As New FileSystemObject

For i = 0 To UBound(aString)
    cName = (i + 1)
    cFile = App.Path & "\txt\" & cName & ".txt"
    
    On Error GoTo myerror
    Open cFile For Output As #1   ' Open file.
    Print #1, aString(i)        ' Read line into variable.
    MsgBox "DONE " & "(" & (i + 1) & ")"
    Close #1   ' Close file.
Next
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function getTableString() As String
'Dim cString As String
'cString = "SELECT FILE1_10.CODE,FILE1_10.DESCA,FILE1_10.MOBIL" & _
'           " From File1_10 INNER JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE  " & _
'           " AND FILE6_20H.doc_no = dbo.f_meeting_doc(FILE1_10.CODE," & cmdYear(0).Tag & ")"
'cWhere = "FILE1_10.Date_Begin <= " & DateSq(DateAdd("yyyy", -1, xDate.text))
'cWhere = cWhere & turn(cWhere, " AND ") & "( NOT FILE1_10.MOBIL IS NULL) AND LEN(FILE1_10.MOBIL) >= 11"
'
'If IsDate(xdate1.text) Then
'    cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.DATE >= " & DateSq(xdate1.text)
'End If
'
'If IsDate(xdate2.text) Then
'    cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.DATE <= " & DateSq(xdate2.text)
'End If
'
'cWhere2 = cWhere & " AND " & "FILE1_10.DIED = 0"
'If cWhere2 <> "" Then cString = cString & " WHERE " & cWhere2
'cString = cString & " ORDER BY FILE1_10.CODE"
'getTableString = cString
End Function
