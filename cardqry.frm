VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cardQryfrm 
   Caption         =   "ÿ»«⁄… «·ﬂ«—‰ÌÂ« "
   ClientHeight    =   9180
   ClientLeft      =   -2415
   ClientTop       =   2415
   ClientWidth     =   18525
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
   ScaleHeight     =   9180
   ScaleWidth      =   18525
   WindowState     =   2  'Maximized
   Begin VB.CheckBox xDamage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… »œ· ›«ﬁœ Ê»œ·  «·›"
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
      Left            =   4995
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   990
      Width           =   2265
   End
   Begin VB.CheckBox xPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·„”œœÌ‰ ›ﬁÿ"
      Enabled         =   0   'False
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
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   990
      Width           =   2265
   End
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·–Ì ·„ Ìÿ»⁄ ›ﬁÿ"
      Enabled         =   0   'False
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
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   990
      Width           =   2265
   End
   Begin VB.Frame Frame9 
      Caption         =   " Õﬁﬁ „‰ «·ﬂ«—‰Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   765
      Width           =   2760
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   41
         Top             =   315
         Width           =   2580
      End
      Begin VB.Label xUnCode 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H001111AE&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   765
         Width           =   2580
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "„Ê”„ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   1215
      Width           =   1905
      Begin VB.TextBox xSeason 
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
         Height          =   390
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Text            =   "2015"
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   1260
      Width           =   6855
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "cardqry.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1680
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   3405
         Picture         =   "cardqry.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   135
         Width           =   1680
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   555
         Left            =   5130
         TabIndex        =   34
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cardqry.frx":4C15
         Caption         =   "√÷«›… "
         Alignment       =   4
      End
      Begin Threed.SSCommand cmdprintrep 
         Height          =   555
         Left            =   1710
         TabIndex        =   35
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cardqry.frx":7C7E
         Caption         =   "ÿ»«⁄…  ﬁ—Ì—"
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   8805
      Width           =   18525
      _ExtentX        =   32676
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10485
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8010
      Width           =   3030
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   1
         Left            =   2655
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»œÊ‰ ’Ê—…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   315
         Width           =   915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   0
         Left            =   1170
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " „  ÿ»«⁄ Â"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   13860
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   4605
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry.frx":A282
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry.frx":CB1C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdSavePrint 
         Caption         =   " „  «·ÿ»«⁄…"
         Height          =   390
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton cmdLastFillgrd 
         Caption         =   "«” —Ã«⁄ «Œ— ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   135
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   12285
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   6135
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   855
         Width           =   1545
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   855
         Width           =   1545
      End
      Begin VB.TextBox xAppend 
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
         TabIndex        =   2
         Top             =   495
         Width           =   960
      End
      Begin VB.TextBox xCode2 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   495
         Width           =   960
      End
      Begin VB.TextBox xCode1 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "≈·Ì"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   540
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   " ⁄÷Ê  «»⁄ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   990
      End
      Begin VB.Label xcode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   135
         Width           =   3840
      End
      Begin VB.Label Label1 
         Caption         =   "«·ﬂÊœ"
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   690
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5955
      Left            =   90
      TabIndex        =   5
      Top             =   2025
      Width           =   18420
      _cx             =   32491
      _cy             =   10504
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
      Rows            =   1
      Cols            =   11
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
      Caption         =   "÷»ÿ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   15885
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7965
      Width           =   2490
      Begin VB.TextBox xDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   315
         Width           =   570
      End
      Begin VB.TextBox xRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   315
         Width           =   435
      End
      Begin VB.Label Label9 
         Caption         =   "Ì„Ì‰ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
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
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "«”›· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   990
      Top             =   720
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
      Left            =   0
      Top             =   675
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
      Left            =   135
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
      Left            =   2430
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
   Begin VB.Frame Frame2 
      Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   13545
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   7965
      Width           =   2310
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label7 
         Caption         =   "«·⁄„Êœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "«·’› :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   420
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   420
      Left            =   0
      Top             =   0
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
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   7965
      Width           =   5100
      Begin ComctlLib.ProgressBar prog1 
         Height          =   555
         Left            =   45
         TabIndex        =   28
         Top             =   180
         Visible         =   0   'False
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   979
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "CardQryfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFilesave As String, cFilePrint As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim printTable As New ADODB.Recordset
Private Sub CmdAdd_Click()
'checkErr
myloadgrd
'On Error Resume Next
grid1.SaveGrid cFilesave, flexFileData
Err.Clear
cmdPrint.Enabled = (grid1.Rows > 1)
checkPhoto
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·ﬂ· !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
    grid1.Rows = 1
    grid1.SaveGrid cFilesave, flexFileData
'    DefineText Me
    Calctotals
End If
End Sub

Private Sub cmdExel_Click()
For i = 1 To grid1.Rows - 1
    If Not validPhoto(RetPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.RowHidden(i) = True
    End If
Next
ToFileExel grid1
For i = 1 To grid1.Rows - 1
    grid1.RowHidden(i) = False
Next
End Sub

Private Sub CmdPrint_Click()
'If Not doPrintMaster Then
If grid1.Rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
    Exit Sub
End If
If Val(xRow.Text) > 5 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 2 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
'If Not doPrintMaster Then
If Not doprint Then
    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
    Exit Sub
End If
Set CardPrintNew.myForm = Me
CardPrintNew.PrintArray
CardPrintNew.Show 1
If MsgBox(" „  «·ÿ»«⁄…", vbYesNo) = vbYes Then
    SavePrinted
    checkPhoto
End If
End Sub
Private Sub CmdExit_Click()
Unload Me
Set cardqry = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdClear_Click()
grid1.Rows = 1
End Sub
Private Sub cmdMember_Click()
Member.Show 1
End Sub
Private Sub CmdLastFillGrd_Click()
Dim fs As New FileSystemObject
If fs.FileExists(cFilesave) Then
    grid1.LoadGrid cFilesave, flexFileData
    If grid1.Rows > 1 Then cmdPrint.Enabled = True
    checkPhoto
End If
End Sub

Private Sub Command1_Click()
If grid1.Rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
    Exit Sub
End If
If Val(xRow.Text) > 5 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 2 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
If Not doprint Then
    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
    Exit Sub
End If
Set CardPrintNew.myForm = Me
CardPrintNew.PrintArray
CardPrintNew.Show 1
If MsgBox(" „  «·ÿ»«⁄…", vbYesNo) = vbYes Then
    SavePrinted
    checkPhoto
End If
End Sub

Private Sub cmdprintrep_Click()
Set PrintGrdNew.myForm = Me
Dim i As Long
For i = 1 To grid1.Rows - 1
    If Not validPhoto(RetPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.RowHidden(i) = True
    End If
Next
PrintGrdNew.doprint grid1, 1, -3, "ÿ»«⁄… «·ÿ·»…", , , , False, False, 9, , aRow
PrintGrdNew.Show 1
For i = 1 To grid1.Rows - 1
    grid1.RowHidden(i) = False
Next
End Sub
Private Sub Form_Load()
MFocus Me
openCon con
cFilesave = App.Path & "\" & Me.Name & ".grd"
Fixgrd
LoadText Me
xSeason.Text = "2015"
xPrinted.Value = 1
xPaid.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
grid1.SaveGrid App.Path & "\cardPrint\Member.grd", flexFileData
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
grid1.SaveGrid cFilesave, flexFileData
End Sub

Private Sub Grid1_EnterCell()
grid1.Editable = IIf(grid1.Col = 5, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
With grid1
If .Rows = 1 Then Exit Sub
If KeyCode = 46 Then
    .RemoveItem grid1.Row
    .SaveGrid cFilesave, flexFileData
End If
End With
End Sub
Private Sub xCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    xUnCode.Caption = ""
    If xCode.Text = "" Then Exit Sub
    xUnCode.ForeColor = -2147483630
    If Val(unMyCodeBar(xCode.Text, "1")) <> 1 Then
        xUnCode.Caption = "Error"
        xUnCode.ForeColor = vbRed
    Else
        xUnCode.Caption = unMyCodeBar(xCode.Text)
    End If
    myGotFocus xCode
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub

Private Sub xCode1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAdd_Click
End If
End Sub

Private Sub xCode1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim cWhere As String
    MemberLookupAll Me, oSearch
End If
End Sub

Private Sub xCode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAdd_Click
End If
End Sub

Private Sub xCODE2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub
Private Function CountGrid() As Integer
With grid1
For i = 1 To grid1.Rows - 1
    'If .TextMatrix(I, 6) = True Then CountGrid = CountGrid + 1
    CountGrid = CountGrid + 1
Next
End With
End Function
Private Sub countPrint()
nCountPrint = 0
With grid1
For i = 1 To .Rows - 1
   If .TextMatrix(i, 6) = True Then nCountPrint = nCountPrint + 1
Next
lblCount.Caption = nCountPrint / 1
End With
End Sub
Private Function MakeString()
MakeString = "#" & ";"
MakeString = MakeString & "|#" & 0 & ";" & "ﬂ«—‰ÌÂ ÃœÌœ"
MakeString = MakeString & "|#" & 1 & ";" & "»œ· ›«ﬁœ"
End Function
Private Sub SavePrinted()
With grid1
'Screen.MousePointer = 11
dTime = Time
dDate = Date
Dim aInsert As Variant
con.BeginTrans
For i = 1 To .Rows - 1
   If validPhoto(RetPhoto(grid1.TextMatrix(i, 2))) Then
        aInsert = AddFlag(Empty, "MEMBER", grid1.TextMatrix(i, 0))
        aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DESCA2", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "RELATION", addstring(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "DATEBIRTH", addDate(grid1.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "DATE", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        aInsert = AddFlag(aInsert, "YEAR", addvalue(xSeason.Text))
        con.Execute addInsert(aInsert, "file4_10")
    End If
Next
con.CommitTrans
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
End Sub
Function eofGrd(cId) As Boolean
eofGrd = (grid1.FindRow(cId, , 0) = -1)
End Function
Private Function doprint() As Boolean
SettingArray(cUpMargin) = MyMeasure(1.7) + MyMeasure(Val(xDown.Text) / 10)
SettingArray(cRightMargin) = MyMeasure(1) + MyMeasure(Val(xRight.Text) / 10)
SettingArray(cCardWidth) = MyMeasure(9.5)
SettingArray(cCardHeight) = MyMeasure(5.81)
SettingArray(cRows) = 5
SettingArray(cCols) = 2
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
nRows = SettingArray(cRows)
nSpace = 0.55
nUP = 0.3

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * nCols) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > nRows, 1, nRow)
    blastrow = (nRow = nRows)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡


prog1.Value = 0
prog1.Visible = True

For i = 1 To grid1.Rows - 1
    If validPhoto(RetPhoto(grid1.TextMatrix(i, 2))) Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 1.1
        
        ' „Â‰œ”
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.2) - MyMeasure(0.1)
        tCard!Top = IIf(grid1.TextMatrix(i, 1) = "", MyMeasure(1.15), MyMeasure(1.1)) - MyMeasure(nDiffer)
        tCard!Width = MyMeasure(1.5)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!fontsize = 11
        tCard!Text = ": „Â‰œ”"
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        
        tCard.AddNew
        tCard!Right = MyMeasure(1.35) + MyMeasure(nDiffer) - MyMeasure(0.45)
        tCard!Top = IIf(grid1.TextMatrix(i, 1) = "", MyMeasure(1.15), MyMeasure(1.1)) - MyMeasure(nDiffer)
        tCard!Width = MyMeasure(5.5)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!fontsize = 11
        tCard!Text = TurnValue(grid1.TextMatrix(i, 3), "", Null)
        tCard!CardNo = nCard
        tCard.Update
        
        ' —ﬁ„ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.1)
        tCard!Top = IIf(grid1.TextMatrix(i, 1) = "", MyMeasure(1.7), MyMeasure(1.1)) + MyMeasure(nSpace) - MyMeasure(nDiffer)
        tCard!Width = MyMeasure(2.1)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!fontsize = 11
        tCard!Text = ": ⁄÷ÊÌ‹… —ﬁ„"
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        
        tCard.AddNew
        tCard!Right = MyMeasure(1.35) + MyMeasure(nDiffer) + MyMeasure(0.3)
        tCard!Top = IIf(grid1.TextMatrix(i, 1) = "", MyMeasure(1.7), MyMeasure(1.1)) + MyMeasure(nSpace) - MyMeasure(nDiffer)
        tCard!Width = 0
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!fontsize = 11
        tCard!Text = ArbString(grid1.TextMatrix(i, 1) & IIf(grid1.TextMatrix(i, 1) <> "", "-", "") & grid1.TextMatrix(i, 0))
        tCard!TextAlign = taLeftTop
        tCard!CardNo = nCard
        tCard.Update
       
        ' «·„—«›ﬁ
        If grid1.TextMatrix(i, 1) <> "" Then
            tCard.AddNew
            tCard!Right = MyMeasure(0.5) + MyMeasure(0.1)
            tCard!Top = MyMeasure(1.1) + MyMeasure(nSpace * 2) - MyMeasure(nDiffer)
            tCard!Width = MyMeasure(1.5)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 11
            tCard!Text = ": «·„—«›ﬁ"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(2)
            tCard!Top = MyMeasure(1.1) + MyMeasure(nSpace * 2) - MyMeasure(nDiffer)
            tCard!Width = MyMeasure(5)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 11
            tCard!Text = TurnValue(grid1.TextMatrix(i, 4), "", Null)
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
       
            ' ’·… «·ﬁ—«»…
            tCard.AddNew
            tCard!Right = MyMeasure(0.5) + MyMeasure(0.2)
            tCard!Top = MyMeasure(1.1) + MyMeasure(nSpace * 3) - MyMeasure(nDiffer)
            tCard!Width = MyMeasure(2.1)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 12
            tCard!Text = " : ’·… «·ﬁ—«»…"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(2.45) + MyMeasure(0.3)
            tCard!Top = MyMeasure(1.1) + MyMeasure(nSpace * 3) - MyMeasure(nDiffer)
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 11
            tCard!Text = TurnValue(grid1.TextMatrix(i, 5), "", Null)
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
        End If
        
        ' «·ﬂ«—‰ÌÂ ‘Œ’Ì
        tCard.AddNew
        tCard!Right = MyMeasure(0.6)
        tCard!Top = MyMeasure(1.3) + MyMeasure(nSpace * 4) + MyMeasure(0.1) - MyMeasure(nDiffer)
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbRed
        tCard!fontsize = 8
        'tCard!Text = "Â–« «·ﬂ«—‰ÌÂ ‘Œ’Ì Ìﬁœ„ ⁄‰œ ﬂ· ÿ·»"
        tCard!Width = MyMeasure(4.8)
        tCard!Height = MyMeasure(0.6)
        tCard!ISBARCODE = True
        tCard!Text = MyCodeBar(grid1.TextMatrix(i, 2), "1")
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
       
        ' «·ﬂ«—‰ÌÂ ‘Œ’Ì
        tCard.AddNew
        tCard!Right = MyMeasure(0.8) + MyMeasure(0.8)
        tCard!Top = MyMeasure(1.5) + MyMeasure(nSpace * 4) + MyMeasure(0.1) - MyMeasure(nDiffer)
        tCard!Width = MyMeasure(3.8)
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbRed
        tCard!fontsize = 8
        'tCard!Text = "ÊÌ”Õ» ›Ï Õ«·… «⁄«— Â ··€Ì—"
        tCard!TextAlign = taCenterTop
        tCard!CardNo = nCard
        tCard.Update
       
'        tCard.AddNew
'        tCard!Right = MyMeasure(0.5)
'        tCard!Top = MyMeasure(1.1) + MyMeasure(0.65 * 4) + MyMeasure(0.4) - MyMeasure(0.1)
'        tCard!Width = 0
'        tCard!Height = 0
'        tCard!FontName = "simplified arabic"
'        tCard!FontBold = True
'        tCard!ForeColor = &HFF&
'        tCard!FontSize = 10
'        tCard!Text = "Ê·« Ì”„Õ »≈⁄«— Â ··€Ì—"
'        tCard!TextAlign = taRightTop
'        tCard!CardNo = nCard
'        tCard.Update
       
        ' Ì‰ ÂÌ ›Ì
        tCard.AddNew
        tCard!Right = MyMeasure(2.5) + MyMeasure(0.3)
        tCard!Top = MyMeasure(1.4) + MyMeasure(nSpace * 4) + MyMeasure(0.7) - MyMeasure(nDiffer)
        tCard!Width = 0
        tCard!Height = 0
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!fontsize = 9
        tCard!Text = "Ì‰ ÂÌ ›Ì " & Format("31/12/2015", "yyyy/m/d")
        tCard!CardNo = nCard
        tCard.Update
        
' «·Ã“¡ «·«Ì”—
        
        'ﬂ·„… —∆Ì” ‰ﬁ«»… «·«”ﬂ‰œ—Ì…
        tCard.AddNew
        tCard!Right = MyMeasure(6.6) - MyMeasure(0.2)
        tCard!Top = MyMeasure(2.8) - MyMeasure(0.1) - MyMeasure(nDiffer) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Width = MyMeasure(3)
        tCard!Height = 1000
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = &HFF&
        tCard!fontsize = 9
        tCard!TextAlign = taCenterTop
        tCard!Text = "‰ﬁÌ» „Â‰œ”Ï «·«”ﬂ‰œ—Ì…"
        tCard!CardNo = nCard
        tCard.Update
                
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.5) - MyMeasure(0.2) + MyMeasure(0.3)
        tCard!Top = MyMeasure(0.62) - MyMeasure(nDiffer) + MyMeasure(0.3)
        tCard!Width = MyMeasure(2.4) * 0.8
        tCard!Height = MyMeasure(2.8) * 0.8
        tCard!Text = RetPhoto(grid1.TextMatrix(i, 6))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        
'        If validPhoto(RetPhoto(grid1.TextMatrix(I, 7)) & "") Then
'            '«·’Ê—… «·’€Ì—…
'            tCard.AddNew
'            tCard!Right = MyMeasure(5.4) - MyMeasure(0.2) + MyMeasure(0.3)
'            tCard!Top = MyMeasure(2.9) - MyMeasure(1.15) - MyMeasure(nDiffer) + MyMeasure(0.3)
'            tCard!Width = MyMeasure(1)
'            tCard!Height = MyMeasure(1.1)
'            tCard!Text = TurnValue(RetPhoto(grid1.TextMatrix(I, 7)), "", Null)
'            tCard!isPhoto = True
'            tCard!CardNo = nCard
'            tCard.Update
'        End If
'        '«· ÊﬁÌ⁄
        tCard.AddNew
        tCard!Right = MyMeasure(6.7) - MyMeasure(0.3) + MyMeasure(0.3)
        tCard!Top = MyMeasure(3.4) - MyMeasure(0.2) - MyMeasure(nDiffer) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Width = MyMeasure(1.9)
        tCard!Height = MyMeasure(0.9)
        tCard!Text = TurnValue(App.Path & "\sign.jpg", "", Null)
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        
        '«”„ —∆Ì” ‰ﬁ«»… «·«”ﬂ‰œ—Ì…
        tCard.AddNew
        tCard!Right = MyMeasure(6.9) - MyMeasure(0.8)
        tCard!Top = MyMeasure(4.4) - MyMeasure(0.5) - MyMeasure(nDiffer) + MyMeasure(0.3)
        tCard!Width = MyMeasure(3.1)
        tCard!Height = 1000
        tCard!FontName = "simplified arabic"
        tCard!FontBold = True
        tCard!ForeColor = &H800000
        tCard!fontsize = 9
        tCard!TextAlign = taCenterTop
        tCard!Text = "„/”„— ‘·»Ì"
        tCard!CardNo = nCard
        tCard.Update
    
    End If
Next
prog1.Visible = False
tCard.Requery
doprint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Private Function doPrintMaster(Optional bSign As Boolean = False) As Boolean
SettingArray(cUpMargin) = MyMeasure(0) + MyMeasure(Val(xDown.Text) / 100)
SettingArray(cRightMargin) = MyMeasure(1.2) + MyMeasure(Val(xRight.Text) / 100)
SettingArray(cCardWidth) = MyMeasure(9.65)
SettingArray(cCardHeight) = MyMeasure(5.8)
SettingArray(cRows) = 5
SettingArray(cCols) = 2
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
nRows = SettingArray(cRows)
nUP = 1.2

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * nCols) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > nRows, 1, nRow)
    blastrow = (nRow = nRows)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡

prog1.Value = 0
prog1.Visible = True

For i = 1 To grid1.Rows - 1
    If validPhoto(RetPhoto(grid1.TextMatrix(i, 0))) Then
        nSpace = MyMeasure(1.5)
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nLine = 0
        prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
           
        If Not bSign Then
            ' «·«”„
            tCard.AddNew
            tCard!Right = MyMeasure(0.8)
            tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
            tCard!Width = MyMeasure(6.2)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 12
            tCard!Text = ArbString("«·«”„ : " & grid1.TextMatrix(i, 1))
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            nLine = nLine + 1
            
            '  «—ÌŒ «·„Ì·«œ
            tCard.AddNew
            tCard!Right = MyMeasure(0.8)
            tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
            tCard!Width = MyMeasure(8)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 12
            tCard!Text = ArbString("«· Œ’’ : " & grid1.TextMatrix(i, 2))
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            nLine = nLine + 1
            tCard.AddNew
            tCard!Right = MyMeasure(0.8)
            tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 13
            tCard!Text = "«·⁄«„ «·œ—«”Ì 2014/2013"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            nLine = nLine + 1
            
    '        nLine = nLine + 1
    '        tCard.AddNew
    '        tCard!Right = MyMeasure(0.8)
    '        tCard!Top = MyMeasure(nUP) + (nSpace * nLine)
    '        tCard!Width = 0
    '        tCard!Height = 0
    '        tCard!FontName = "simplified arabic"
    '        tCard!FontBold = True
    '        tCard!ForeColor = vbBlack
    '        tCard!FontSize = 13
    '        tCard!Text = ArbString("«·—ﬁ„ «·ﬁÊ„Ì : " & grid1.TextMatrix(I, 4))
    '        tCard!TextAlign = taRightTop
    '        tCard!CardNo = nCard
    '        tCard.Update
    '        nLine = nLine + 1
               
            '«·’Ê—… «·ﬂ»Ì—…
            tCard.AddNew
            tCard!Right = MyMeasure(6.5)
            tCard!Top = MyMeasure(1)
            tCard!Width = MyMeasure(2.1)
            tCard!Height = MyMeasure(2.8)
            tCard!Text = RetPhoto(grid1.TextMatrix(i, 0))
            tCard!isPhoto = True
            tCard!CardNo = nCard
            tCard.Update
            
            ' ”ﬂ— Ì— ⁄«„ «·›—⁄
            tCard.AddNew
            tCard!Right = MyMeasure(6.3)
            tCard!Top = MyMeasure(3.65)
            tCard!Width = MyMeasure(2.5)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbRed
            tCard!fontsize = 11
            tCard!Text = "⁄„Ìœ «·ﬂ·Ì…"
            tCard!TextAlign = taCenterTop
            tCard!CardNo = nCard
            tCard.Update
            nLine = nLine + 1
            
            ' «”„ ”ﬂ— Ì— ›—⁄ «·«”ﬂ‰œ—Ì…
            tCard.AddNew
            tCard!Right = MyMeasure(6.3)
            tCard!Top = MyMeasure(4.5)
            tCard!Width = MyMeasure(2.5)
            tCard!Height = 0
            tCard!FontName = "simplified arabic"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!fontsize = 11
            tCard!Text = "√.œ/”„Ì— ﬂ«„·"
            tCard!TextAlign = taCenterTop
            tCard!CardNo = nCard
            tCard.Update
                
                '«· ÊﬁÌ⁄
            tCard.AddNew
            tCard!Right = MyMeasure(6.3)
            tCard!Top = MyMeasure(4.2)
            tCard!Width = MyMeasure(2.24)
            tCard!Height = MyMeasure(0.51)
            tCard!Text = App.Path & "\sign.jpg"
            tCard!isPhoto = True
            tCard!CardNo = nCard
            tCard.Update
        Else
            tCard.AddNew
            tCard!Right = MyMeasure(5)
            tCard!Top = MyMeasure(2.4)
            tCard!Width = MyMeasure(2)
            tCard!Height = MyMeasure(2)
            tCard!Text = App.Path & "\nesr.jpg"
            tCard!isPhoto = True
            tCard!CardNo = nCard
            tCard.Update
        End If
    End If
Next
prog1.Visible = False
tCard.Requery
doPrintMaster = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function

Sub myProc()
ActiveControl.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Private Sub checkPhoto()
Dim aPrint As Variant
With grid1
prog1.Value = 0
prog1.Visible = True
For i = 1 To grid1.Rows - 1
    prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
    If Not validPhoto(RetPhoto(grid1.TextMatrix(i, 6))) Then grid1.Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
    aPrint = Printed(.TextMatrix(i, 2), xSeason.Text, con)
    grid1.TextMatrix(i, .Cols - 1) = Format(retFlag(aPrint, "date"), "yyyy/mm/dd")
    If grid1.TextMatrix(i, .Cols - 1) <> "" Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE0E0E0
Next
prog1.Visible = False
End With
End Sub
Private Sub xCONTEST_Click(Area As Integer)
xCONTEST_Validate False
End Sub
Private Sub xCONTEST_Validate(Cancel As Boolean)
If Not xContest.MatchedWithList Then xContest.BoundText = ""
Dim cString As String
cString = "Select code,desca from age_codes"
If IsNumeric(xContest.BoundText) Then cString = cString & turn(cString) & " contest = " & xContest.BoundText
If DATA3.RecordSource <> cString Then
    DATA3.RecordSource = cString
    cString = xAge.BoundText
    DATA3.Refresh
    xAge.BoundText = cString
    If Not xAge.MatchedWithList Then xAge.BoundText = ""
End If
End Sub

Private Sub xDEGREE_Validate(Cancel As Boolean)
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
Dim cString As String
cString = "Select code,desca from class_codes"
If xDegree.MatchedWithList Then cString = cString & turn(cString) & " degree = " & xDegree.BoundText
cString = cString & " order by code"
FixClass
End Sub

Private Sub xDown_Change()
' addSetting "down", Val(xDown.Text), cFilePrint
End Sub

Private Sub xDpt_Click(Area As Integer)
FixClass
End Sub

Private Sub xLang_Validate(Cancel As Boolean)
FixClass
End Sub

Private Sub xReg_Validate(Cancel As Boolean)
FixClass
End Sub

Private Sub xSeason_Validate(Cancel As Boolean)
'CmdPrint.Enabled = (Val(xSeason.Text) = 2010)

End Sub
Private Sub myloadgrd()
Dim loctable As New ADODB.Recordset, aDamage As Variant
loctable.Open "select * from file1_20 where isCard = 1 or isDamage = 1", con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    aDamage = AddFlag(aDamage, loctable!CODE, "ok")
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing

Dim GRDTABLE As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, i As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT FILE1_10.* FROM FILE1_10 LEFT JOIN FILE6_20H ON FILE1_10.CODE = FILE6_20H.CODE"
If ValidInt(xCode1.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate2.Text)
End If

If ValidInt(xCode2.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.Text
End If

cString = cString & " ORDER BY FILE1_10.DESCA"

Set GRDTABLE = New ADODB.Recordset
With grid1
GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (GRDTABLE.EOF And GRDTABLE.BOF) Then
    GRDTABLE.MoveLast
    nRecordcount = GRDTABLE.RecordCount
    GRDTABLE.MoveFirst
End If
prog1.Visible = True
prog1.Value = 0
Do Until GRDTABLE.EOF
    i = i + 1
    bAddRow = .FindRow(GRDTABLE!CODE, , 0) = -1
    
    aPrint = Printed(GRDTABLE!CODE, xSeason.Text, con)
    aPaid = LastDoc_card(GRDTABLE!CODE, con)
    
    If bAddRow And xPrinted.Value = 1 Then
        bAddRow = IsEmpty(aPrint)
        If (Not IsEmpty(aPaid)) And (Not IsEmpty(aPrint)) And xDamage.Value = 1 Then
            If retFlag(aDamage, retFlag(aPaid, "ITEM")) = "ok" Then
                bAddRow = True
            End If
        End If
    End If
    
    If bAddRow And xPaid.Value = 1 Then
        bAddRow = Not IsEmpty(aPaid)
    End If
        
    If bAddRow Then
        prog1.Value = Round(i / (nRecordcount), 2) * 100
        
        .AddItem ""
        .TextMatrix(grid1.Rows - 1, 0) = GRDTABLE!CODE
        .TextMatrix(grid1.Rows - 1, 1) = ""
                                            
        ' «·„⁄—Ê÷
        .TextMatrix(grid1.Rows - 1, 2) = GRDTABLE!CODE
        .TextMatrix(grid1.Rows - 1, 3) = GRDTABLE!desca
        .TextMatrix(grid1.Rows - 1, 4) = ""
        .TextMatrix(grid1.Rows - 1, 5) = ""
        
        ' «·’Ê—
        .TextMatrix(grid1.Rows - 1, 6) = GRDTABLE!CODE
        .TextMatrix(grid1.Rows - 1, 7) = GRDTABLE!CODE
        .TextMatrix(grid1.Rows - 1, 8) = Format(retFlag(aPaid, "doc_no"))
        .TextMatrix(grid1.Rows - 1, 9) = Format(GRDTABLE!dateBirth, "DD-MM-YYYY")
        .TextMatrix(grid1.Rows - 1, 10) = Format(retFlag(aPrint, "date"), "YYYY/M/D")
    End If
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing

cString = "SELECT file1_11.Code,File1_11.Relation,File1_11.Member,File1_11.Title,file1_11.descA,File1_11.DOC_NO,FILE1_11.DATEBIRTH,FILE1_10.[UNION_REG],FILE1_10.DESCA as Desca_Member,Relation_codes.Desca as Desca_relation " & _
        " From File1_11 Inner Join File1_10 On File1_11.Member = File1_10.Code INNER join relation_codes on file1_11.relation = relation_codes.code LEFT JOIN FILE6_20H ON FILE6_20H.CODE = FILE1_11.MEMBER"

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate2.Text)
End If


If xAppend.Text = "" Then
    If ValidInt(xCode1.Text) Then
        cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
    End If
    
    If ValidInt(xCode2.Text) Then
        cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.Text
    End If
Else
    cString = cString & turn(cString) & "File1_11.MEMBER = " & xCode1.Text & " and File1_11.code = " & xAppend.Text
End If


Set GRDTABLE = New ADODB.Recordset
GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (GRDTABLE.EOF And GRDTABLE.BOF) Then
    GRDTABLE.MoveLast
    nRecordcount = GRDTABLE.RecordCount
    GRDTABLE.MoveFirst
End If
prog1.Visible = True
prog1.Value = 0


Do Until GRDTABLE.EOF
    i = i + 1
    bAddRow = .FindRow(GRDTABLE!Member & "-" & GRDTABLE!CODE, , 2) = -1
    aPrint = Printed(GRDTABLE!Member & "-" & GRDTABLE!CODE, xSeason.Text, con)
    aPaid = LastDoc_card(GRDTABLE!Member, con, GRDTABLE!CODE)
    
    If bAddRow And xPrinted.Value = 1 Then
        bAddRow = IsEmpty(aPrint)
        If (Not IsEmpty(aPaid)) And (Not IsEmpty(aPrint)) And xDamage.Value = 1 Then
            If retFlag(aDamage, retFlag(aPaid, "ITEM")) = "ok" Then
                bAddRow = True
            End If
        End If
    End If
    
    If bAddRow And xPaid.Value = 1 Then
        bAddRow = Not IsEmpty(aPaid)
    End If
    
    If bAddRow Then
         prog1.Value = IIf(Round(i / (nRecordcount), 2) > 1, 1, Round(i / (nRecordcount), 2)) * 100
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = GRDTABLE!Member
        .TextMatrix(.Rows - 1, 1) = GRDTABLE!CODE
         
         ' «·„⁄—Ê÷
        .TextMatrix(.Rows - 1, 2) = GRDTABLE!Member & "-" & GRDTABLE!CODE
        .TextMatrix(.Rows - 1, 3) = GRDTABLE!Desca_Member
        .TextMatrix(.Rows - 1, 4) = GRDTABLE!desca & ""
        .TextMatrix(.Rows - 1, 5) = GRDTABLE!Desca_relation
         
         ' «·’Ê—
        .TextMatrix(.Rows - 1, 6) = GRDTABLE!Member & "-" & GRDTABLE!CODE
        .TextMatrix(.Rows - 1, 7) = GRDTABLE!Member
        .TextMatrix(grid1.Rows - 1, 8) = Format(retFlag(aPaid, "doc_no"))
        .TextMatrix(grid1.Rows - 1, 9) = Format(GRDTABLE!dateBirth, "DD-MM-YYYY")
        .TextMatrix(grid1.Rows - 1, 10) = Format(retFlag(aPrint, "date"), "YYYY/M/D")
    End If
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing
prog1.Visible = False
Me.MousePointer = 0
If grid1.Rows > 1 Then grid1.Select 1, 0, 1, 1
grid1.Sort = flexSortGenericAscending
Calctotals
End With
End Sub
Private Sub Fixgrd()
With grid1
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷Ê"
    .TextMatrix(0, 3) = "≈”„ «·„Â‰œ”"
    .TextMatrix(0, 4) = "«·„—«›ﬁ"
    .TextMatrix(0, 5) = "œ—Ã… «·ﬁ—«»…"
    .TextMatrix(0, 8) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 9) = " «—ÌŒ «·„Ì·«œ"
    .TextMatrix(0, 10) = " «—ÌŒ «·ÿ»«⁄…"
    
    .ColHidden(0) = True
    .ColHidden(1) = True
    .ColHidden(6) = True
    .ColHidden(7) = True
    .ColHidden(9) = True
    
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1500
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColDataType(0) = flexDTLong
    .ColDataType(1) = flexDTLong

End With
End Sub
Private Sub FixClass()
Dim cString As String, cSave As String
cString = "Select * from class_codes"
If xDegree.MatchedWithList Then cString = cString & turn(cString) & "[degree] = " & xDegree.BoundText
If xType.MatchedWithList Then cString = cString & turn(cString) & "[type] = " & xType.BoundText
If xLang.MatchedWithList Then cString = cString & turn(cString) & "[lang] = " & xLang.BoundText
If xReg.MatchedWithList Then cString = cString & turn(cString) & "[reg] = " & xReg.BoundText
cString = cString & cOrderBy
If Trim(UCase(DATA1.Recordset.Source)) <> Trim(UCase(cString)) Then
    cSave = xClass.BoundText
    Set DATA1.Recordset = myRecordSet(cString, con)
    xClass.BoundText = cSave
    If Not xClass.MatchedWithList Then xClass.BoundText = ""
End If
End Sub

Private Sub xType_Click(Area As Integer)
FixClass
End Sub
Private Sub Calctotals()
Dim nAll As Long, nphoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(1).Text = ""
If grid1.Rows = 1 Then Exit Sub
For i = 0 To grid1.Rows - 1
    nAll = nAll + 1
    If validPhoto(RetPhoto(grid1.TextMatrix(i, 0))) Then nphoto = nphoto + 1
Next
nPhoto2 = ((Val(xRow.Text) - 1) * 2) + (Val(xCol.Text) - 1)
nPages = Fix(nphoto / 10)
If nphoto > 10 Then nLeft = nphoto Mod 10
StatusBar1.Panels(3).Text = "⁄œœ «·”Ã·«  : " & nAll
StatusBar1.Panels(2).Text = "⁄œœ «·”Ã·«  »’Ê— : " & nphoto
StatusBar1.Panels(1).Text = "⁄œœ «·’›Õ«  : " & nPages
If nrest > 0 Then StatusBar1.Panels(3).Text = StatusBar1.Panels(3).Text & turn(StatusBar1.Panels(3).Text, " ") & nrest & " ’Ê—…"
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xSeason_GotFocus()
myGotFocus xSeason
End Sub
Private Sub xSeason_LostFocus()
myLostFocus xSeason
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xAppend_GotFocus()
myGotFocus xAppend
End Sub
Private Sub xAppend_LostFocus()
myLostFocus xAppend
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xcode_desca.Caption = ""
If Not ValidInt(xCode1.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode1.Text)
If Not IsEmpty(aRet) Then
    xcode_desca.Caption = retFlag(aRet, "DESCA") & ""
End If
End Sub
Private Sub xDown_GotFocus()
myGotFocus xDown
End Sub
Private Sub xDown_LostFocus()
myLostFocus xDown
End Sub
Private Sub xRight_GotFocus()
myGotFocus xRight
End Sub
Private Sub xRight_LostFocus()
myLostFocus xRight
End Sub
Private Sub xCol_GotFocus()
myGotFocus xCol
End Sub
Private Sub xCol_LostFocus()
myLostFocus xCol
End Sub
Private Sub xRow_GotFocus()
myGotFocus xRow
End Sub
Private Sub xRow_LostFocus()
myLostFocus xRow
End Sub

