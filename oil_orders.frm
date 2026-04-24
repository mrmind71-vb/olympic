VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form oil_ordersfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” Â·«ﬂ “ÌÊ "
   ClientHeight    =   8295
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   10965
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   10965
   Begin VB.Frame Frame7 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   1215
      Width           =   10770
      Begin VB.TextBox xBon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6840
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "aa"
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·»Ê‰"
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
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   225
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   4950
      Width           =   10770
      Begin VB.TextBox xQuant2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3510
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Tag             =   "aa"
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xGas2 
         Height          =   330
         Left            =   6930
         TabIndex        =   60
         Top             =   180
         Width           =   1950
         _ExtentX        =   3440
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
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
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
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "“Ì  ›·›·Ì‰"
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
         Left            =   8955
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6930
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   2040
   End
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   2835
      Width           =   10770
      Begin VB.TextBox xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   900
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Tag             =   "aa"
         Top             =   900
         Width           =   2175
      End
      Begin VB.TextBox xDistance_car 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6705
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "aa"
         Top             =   945
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xcar_driver 
         Height          =   330
         Left            =   5940
         TabIndex        =   7
         Top             =   585
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
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
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   540
         Width           =   2940
         _ExtentX        =   5186
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
      Begin MSDataListLib.DataCombo xcar 
         Height          =   330
         Left            =   5940
         TabIndex        =   5
         Top             =   225
         Width           =   2940
         _ExtentX        =   5186
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
      Begin MSDataListLib.DataCombo xGas 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   180
         Width           =   2940
         _ExtentX        =   5186
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "«·›—ﬁ"
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
         Left            =   8955
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1710
         Width           =   450
      End
      Begin VB.Label xDiffer 
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
         Height          =   330
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1665
         Width           =   2940
      End
      Begin VB.Label xDistance 
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
         Height          =   330
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1305
         Width           =   2940
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·„”«›… «·›⁄·Ì…"
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·“Ì "
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   225
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "”«∆ﬁ «·”Ì«—…"
         DragMode        =   1  'Automatic
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
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "«·”Ì«—…"
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
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "«·”«∆ﬁ «·„” ⁄„·"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   585
         Width           =   1335
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "«·≈Ã„«·Ì"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   900
         Width           =   660
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ﬂÌ·Ê „ —"
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
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   990
         Width           =   645
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "«·„”«›… «·„› —÷…"
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
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1425
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "„ﬂÊ‰«  «·’‰›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   -2790
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9180
      Visible         =   0   'False
      Width           =   7800
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3300
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   7575
         _cx             =   13361
         _cy             =   5821
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
         Cols            =   6
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
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1845
      Width           =   10770
      Begin VB.TextBox xCode_next 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   225
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   990
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "aa"
         Top             =   180
         Width           =   2040
      End
      Begin VB.TextBox xCounter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6885
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "aa"
         Top             =   540
         Width           =   2040
      End
      Begin VB.TextBox xQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6885
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "aa"
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label xCounter_next 
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
         Height          =   330
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   540
         Width           =   2040
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ—«¡… «· «·Ì…"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·⁄œ«œ"
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   585
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         DragMode        =   1  'Automatic
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1710
         Width           =   60
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   45
      Width           =   6990
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "oil_orders.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2325
         MaskColor       =   &H00FFFFFF&
         Picture         =   "oil_orders.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "oil_orders.frx":48DC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "oil_orders.frx":6D48
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4590
         MaskColor       =   &H00FFFFFF&
         Picture         =   "oil_orders.frx":95E2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5760
         Picture         =   "oil_orders.frx":BB8E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   1080
      Top             =   2070
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   -1215
      Top             =   765
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   -1485
      Top             =   -45
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
   Begin Threed.SSCommand cmdFilter 
      Height          =   285
      Left            =   135
      TabIndex        =   26
      Top             =   7695
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Undo Filter"
      ButtonStyle     =   3
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   -90
      Top             =   450
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
   Begin VB.Frame Frame5 
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   5625
      Width           =   10770
      Begin VB.TextBox xCode_Prv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7605
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   900
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «Œ—  „ÊÌ‰ "
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
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label xDate_Prv 
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
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "›—ﬁ «Œ—  „ÊÌ‰"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   945
         Width           =   1200
      End
      Begin VB.Label xdiffer_Prv 
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label xCounter_Prv 
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
         Height          =   330
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label xDistance_car_Prv 
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label xDistance_Prv 
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "«·„”«›… «·„› —÷…"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "ﬁ—«¡… «Œ—  „ÊÌ‰"
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
         Left            =   9135
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "„”«›… «Œ—  „ÊÌ‰"
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   585
         Width           =   1365
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6930
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "oil_orders.frx":E361
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "oil_orders.frx":10531
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   810
         TabIndex        =   20
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "oil_orders.frx":12679
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "oil_orders.frx":14841
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
         TabIndex        =   21
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "oil_orders.frx":16990
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "oil_orders.frx":18B70
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2340
         TabIndex        =   22
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
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
         Picture         =   "oil_orders.frx":1ACCB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "oil_orders.frx":1CE87
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   855
      Top             =   -45
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„”·”·"
      DragMode        =   1  'Automatic
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
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   855
      Width           =   690
   End
   Begin VB.Label xRecord 
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
      Height          =   465
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7020
      Width           =   4425
   End
End
Attribute VB_Name = "oil_ordersfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bedit As Boolean
Public sBon As String
Dim aEditRecord As Variant
Public sCode As String
Dim nRound As Long
Dim cDef_Gas As String
Dim bActivated As Boolean
Dim con As New ADODB.Connection
Dim cFilter As String, cFilterLookup As String
Dim oSearch As New Search3, oSearchCode As New Search3, oSearchCharge As New Search3
Dim formMode As Byte, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdBons_Click()
'Set bon_addfrm.myForm = Me
'bon_addfrm.nFlag = 2
'bon_addfrm.sBon = xBon.Caption
'bon_addfrm.sNote = xNote.BoundText
'bon_addfrm.sNoteDesca = xNote.Text
'bon_addfrm.Show 1
'If sBon <> "" Then
'    xBon.Caption = sBon
'    xQuant.SetFocus
'    sBon = ""
'End If
End Sub
Private Sub cmdDelFix_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If Not FixOrderDelete Then Exit Sub
    If sCode <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(Xcode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdFilter_Click()
cFilterLookup = ""
openCardTable
myUndo
End Sub
Private Sub cmdGroup_Click()
Dim sBound As String
sBound = xNote.BoundText
notes_codes_oilfrm.Show 1
Set data4.Recordset = myRecordSet("SELECT * FROM NOTES_CODES_OIL", con)
xNote.BoundText = sBound
If Not xNote.MatchedWithList Then xNote.BoundText = ""
End Sub
Private Sub cmdSave_fix_Click()
'If Not MYVALID Then Exit Sub
'If Not FixOrder Then Exit Sub
'Inform " „  ⁄œÌ· «·»Ì«‰«  »‰Ã«Õ"
'openCardTable
'myUndo
End Sub

Private Sub Form_Activate()
If Not bActivated Then
    bActivated = True
    If sCode <> "" Then
        On Error Resume Next
        Xcode.SetFocus
        Err.Clear
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Or TypeOf ActiveControl Is DBCombo) Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con
cDef_Gas = retDef("TYPE_GAS_CODES", "CODE", "KIND = 2")
Set data1.Recordset = myRecordSet("SELECT * FROM CARS", con)
Set xCar.RowSource = data1
xCar.ListField = "Desca"
xCar.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("SELECT * FROM DRIVER", con)
Set xDriver.RowSource = DATA2
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

Set xcar_driver.RowSource = DATA2
xcar_driver.ListField = "Desca"
xcar_driver.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("SELECT * FROM TYPE_GAS_CODES WHERE KIND = 2", con)
Set xGas.RowSource = DATA3
xGas.ListField = "Desca"
xGas.BoundColumn = "Code"


Set DATA5.Recordset = myRecordSet("SELECT * FROM TYPE_GAS_CODES WHERE KIND = 3", con)
Set xGas2.RowSource = DATA5
xGas2.ListField = "Desca"
xGas2.BoundColumn = "Code"

openCardTable
myUndo
If sCode = "" And aEditRecord Then CmdAdd_Click
End Sub
Private Sub CmdAdd_Click()
mydefine
Xcode.Text = ""
On Error Resume Next
xDate.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If Not FixOrderDelete Then Exit Sub
    If sCode <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(Xcode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not FixOrder Then Exit Sub
'If Not myReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If sCode <> "" Then
    Unload Me
    Exit Sub
End If
If Xcode.Tag = DefineMode Then
    CmdAdd_Click
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
xDate.SetFocus
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
Sub Handlecontrols(nMode)
'aEditRecord = bedit And xCode_next.Text = ""
aEditRecord = bedit
cmdSave.Enabled = aEditRecord
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode) And aEditRecord
CmdInform.Enabled = (nMode = LoadMode) And Trim(sCode) = ""
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
cmdFilter.Visible = cFilterLookup <> ""
Xcode.Enabled = Not (nMode = LoadMode)
Xcode.Tag = nMode
End Sub
Sub mydefine()
Xcode.Text = ""
'xBon.caption = Val(xBon.caption) + 1
xBon.Text = ""
xDate.Text = Format(Date, "YYYY-MM-DD")
xDate_Prv.Caption = ""
xCounter_Prv.Caption = ""
xQuant.Text = ""
xCar.BoundText = ""
xcar_driver.BoundText = ""
xDriver.BoundText = ""
xCounter.Text = ""
xCounter_next.Caption = ""
xDistance_car.Text = ""
xDistance.Caption = ""
xDiffer.Caption = ""
xCode_next.Text = ""
xCode_Prv.Text = ""
xDistance_Prv.Caption = ""
xGas.BoundText = cDef_Gas
xGas2.BoundText = cDef_Gas2
xQuant2.Text = ""
xTotal.Text = ""
'xPrice2.Text = ""
xdiffer_Prv.Caption = ""
xDistance_car_Prv.Caption = ""
Handlecontrols DefineMode
xRecord.Caption = "«÷«›… ”Ã· ÃœÌœ"
End Sub
Private Sub myload()
Xcode.Text = CardTable!code & ""
xDate.Text = Format(CardTable!Date, "yyyy-mm-dd")
xDate_Prv.Caption = Format(CardTable!date_prv, "yyyy-mm-dd")
xCounter_Prv.Caption = Myvalue(CardTable!counter_prv)
xQuant.Text = Myvalue(CardTable!Quant)
xCar.BoundText = CardTable!car & ""
xBon.Text = CardTable!BON & ""
xcar_driver.BoundText = CardTable!car_driver & ""
xDriver.BoundText = CardTable!Driver & ""
xCounter.Text = Myvalue(CardTable!Counter)
xCounter_next.Caption = Myvalue(CardTable!counter_next & "")
xDistance_car.Text = Myvalue(CardTable!distance_Car & "")
xDistance.Caption = Myvalue(CardTable!Distance & "")
xDiffer.Caption = Myvalue(CardTable!Differ & "")
xCode_next.Text = CardTable!code_next & ""
xCode_Prv.Text = CardTable!Code_prv & ""
xDistance_Prv.Caption = CardTable!Distance_prv & ""
xGas.BoundText = CardTable!gas & ""
xGas2.BoundText = CardTable!gas2 & ""
'xPrice2.Text = Myvalue(CardTable!Price2)
xQuant2.Text = Myvalue(CardTable!quant2)
xTotal.Text = Myvalue(CardTable!total)
xdiffer_Prv.Caption = Myvalue(CardTable!differ_prv)
xDistance_car_Prv.Caption = Myvalue(CardTable!distance_car_prv)
Calctotals
xRecord.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Calctotals

Dim aPrevious As Variant
aPrevious = retPrevious

Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "Gas", addvalue(xGas.BoundText))
aInsert = AddFlag(aInsert, "Gas2", addvalue(xGas2.BoundText))
aInsert = AddFlag(aInsert, "Quant", Val(xQuant.Text))
aInsert = AddFlag(aInsert, "Quant2", Val(xQuant2.Text))
aInsert = AddFlag(aInsert, "price", Val(xPrice.Text))
aInsert = AddFlag(aInsert, "price2", Val(xPrice2.Text))
aInsert = AddFlag(aInsert, "NOTE", addstring(xNote.BoundText))
aInsert = AddFlag(aInsert, "bon", addstring(xBon.Text))
aInsert = AddFlag(aInsert, "Driver", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "car_Driver", addstring(xcar_driver.BoundText))
aInsert = AddFlag(aInsert, "car", addvalue(xCar.BoundText))
aInsert = AddFlag(aInsert, "Counter", Val(xCounter.Text))
aInsert = AddFlag(aInsert, "Distance_car", Val(xDistance_car.Text))
aInsert = AddFlag(aInsert, "code_prv", addstring(retFlag(aPrevious, "code")))
con.BeginTrans
If Xcode.Text = "" Then
    Xcode.Text = RetZero(Val(Newflag("oil_orders", "code")))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(Xcode.Text))
    con.Execute addInsert(aInsert, "oil_orders")
Else
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(Xcode.Text))
End If

If Not IsEmpty(aPrevious) Then
    aInsert = AddFlag(Empty, "code_next", addstring(Xcode.Text))
    aInsert = AddFlag(aInsert, "Counter_Next", Val(xCounter.Text))
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aPrevious, "code")))
End If
con.CommitTrans

myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    Xcode.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = Xcode.Name Then
    Xcode.Text = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
    SendKeys "{TAB}"
    oSearchCode.Hide
    Unload oSearchCode
End If
End Sub
Sub myproc2(pFilter As String)
Unload oSearchItem
cFilterLookup = pFilter
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload oSearch
Set oSearch = Nothing
Err.Clear
closeCon con
UnloadAllForms "search3"
Unload itemsfrm
Set itemsfrm = Nothing
End Sub

Private Sub XBON_GotFocus()
'If Not xNote.MatchedWithList Then
'    xNote.BoundText = ""
'    XBON.Text = ""
'Else
'    Dim sText As String, bAddText As Boolean
'    sText = XBON.Text
'    XBON.Clear
'    Dim aRet As Variant, cString As String
'    aRet = GetFields("select BON,BON_COUNT FROM NOTES_CODES WHERE CODE = " & xNote.BoundText)
'    If Not IsEmpty(aRet) Then
'        For i = 0 To retFlag(aRet, "BON_COUNT") - 1
'            cString = "select CODE FROM oil_orders WHERE NOTE = " & xNote.BoundText & " AND BON = " & retFlag(aRet, "BON") + i
'            If xCode.Text <> "" Then cString = cString & turn(cString) & "CODE <> " & xCode.Text
'            If IsEmpty(GetField(cString)) Then
'                XBON.AddItem retFlag(aRet, "BON") + i
'            End If
'            If Val(sText) = Val(retFlag(aRet, "BON") + i) Then bAddText = True
'        Next
'    End If
'    If bAddText Then XBON.Text = sText
'End If
End Sub

Private Sub xcar_Click(Area As Integer)
If Area = 2 Then xCar_LostFocus
End Sub
Private Sub xCar_LostFocus()
myLostFocus xCar
If xCar.MatchedWithList Then
    Dim aRet As Variant
    aRet = GetFields("select driver,type_gas from cars where code = " & xCar.BoundText, con)
    xcar_driver.BoundText = retFlag(aRet, "driver") & ""
    xDriver.BoundText = retFlag(aRet, "driver") & ""
    'xGas.BoundText = retFlag(aRet, "type_Gas") & ""
    xGas_LostFocus
'    xGas2_LostFocus
Else
    xcar_driver.BoundText = ""
    xDriver.BoundText = ""
    'xGas.BoundText = ""
    xGas_LostFocus
 '   xGas2_LostFocus
End If
Calctotals
End Sub
Private Sub xCode_LostFocus()
myLostFocus Xcode
If Xcode.Text = "" Then Exit Sub
Xcode.Text = RetZero(Xcode.Text)
CardTable.Find "code = " & MyParn(Xcode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    If Xcode.Tag = LoadMode Then mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not xCar.MatchedWithList Then
    If Not bIgMsg Then MsgBox "»Ì«‰«  «·”Ì«—… €Ì— „”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— „”Ã·"
    Exit Function
End If

If Not xGas.MatchedWithList Then
    If Not bIgMsg Then MsgBox "‰Ê⁄ «·ÊﬁÊœ €Ì— „”Ã·"
    Exit Function
End If


If Not xDriver.MatchedWithList Then
    If Not bIgMsg Then MsgBox "«·”«∆ﬁ €Ì— „”Ã·"
    Exit Function
End If

If Not ValidInt(xBon.Text) Then
    If Not bIgMsg Then MsgBox "—ﬁ„ «·»Ê‰ €Ì— „”Ã·"
    Exit Function
End If

Dim sValid As String
sValid = validPrevious
If sValid <> "ok" Then
    MsgBox ArbString("ﬁ—«¡… ”«»ﬁ… «ﬂ»— „‰ Â–Â «·ﬁ—«¡…" & Space(2) & "(" & sValid & ")")
    Exit Function
End If

sValid = validNext
If sValid <> "ok" Then
    MsgBox ArbString("ﬁ—«¡…  «·Ì… «ﬁ· „‰ Â–Â «·ﬁ—«¡…" & Space(2) & "(" & sValid & ")")
    Exit Function
End If

MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(Xcode.Text) <> "" Then
        CardTable.Find "CODE = " & MyParn(Xcode.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT oil_orders.*,oil_orders_1.COUNTER AS COUNTER_PRV,oil_orders_1.DISTANCE_CAR AS DISTANCE_CAR_PRV," & _
          " oil_orders_1.DATE AS DATE_PRV, oil_orders_1.DIFFER AS DIFFER_PRV,oil_orders_1.DISTANCE AS DISTANCE_PRV " & _
          " FROM oil_orders LEFT JOIN oil_orders AS oil_orders_1 ON oil_orders.CODE_PRV = oil_orders_1.CODE"
cFilter = ""
If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "oil_orders.CODE = " & MyParn(sCode)
If cFilterLookup <> "" Then cFilter = cFilter & turn(cFilter, " and ") & cFilterLookup
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY oil_orders.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub Calctotals()
If Val(xCounter_next.Caption) <> 0 Then
    xDistance.Caption = Myvalue(Int(Val(xCounter_next.Caption) - Val(xCounter.Text)))
    xDiffer.Caption = Myvalue(Val(xDistance.Caption) - Val(xDistance_car.Text))
Else
    xDistance.Caption = ""
    xDiffer.Caption = ""
End If
'xTotal.Text = Myvalue(Val(xQuant.Text) * Val(xPrice.Text))
If Val(xDiffer.Caption) > 0 Then xDiffer.Caption = "+" & xDiffer.Caption
End Sub
Private Sub xCounter_Change()
xCar.Enabled = Val(xCounter.Text) > 0 And ValidInt(xCounter.Text)
End Sub

Private Sub xDate_Change()
xCar.Enabled = IsDate(xDate.Text)
End Sub

Private Sub xDiffer_Change()
If Val(xDiffer.Caption) > 0 Then
    xDiffer.ForeColor = &H8000&
ElseIf Val(xDiffer.Caption) < 0 Then
    xDiffer.ForeColor = vbRed
Else
    xDiffer.ForeColor = vbBlack
End If
End Sub

Private Sub xDistance_car_Change()
Calctotals
'xDiffer.ForeColor = IIf(Val(xDistance_differ.Caption) > 0, vbRed, vbBlack)
'xGas_differ.ForeColor = IIf(Val(xGas_differ.Caption) > 0, vbRed, vbBlack)
End Sub

Private Sub xLine_GotFocus()
myGotFocus xLine
End Sub
Private Sub xLine_LostFocus()
myLostFocus xLine
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xEmp_GotFocus()
myGotFocus xEmp
End Sub
Private Sub xEmp_LostFocus()
myLostFocus xEmp
End Sub
Private Sub xTime_in_GotFocus()
myGotFocus xTime_in
End Sub
Private Sub xCounter_in_GotFocus()
myGotFocus xCounter_in
End Sub
Private Sub xTime_out_GotFocus()
myGotFocus xTime_out
End Sub
Private Sub xCounter_out_GotFocus()
myGotFocus xCounter_out
End Sub
Private Sub xCode_GotFocus()
myGotFocus Xcode
End Sub
Private Sub xcar_driver_GotFocus()
myGotFocus xcar_driver
End Sub
Private Sub xcar_driver_LostFocus()
myLostFocus xcar_driver
If Not xcar_driver.MatchedWithList Then xcar_driver.BoundText = ""
End Sub
Private Sub xDriver_GotFocus()
myGotFocus xDriver
End Sub
Private Sub xDriver_LostFocus()
myLostFocus xDriver
If Not xDriver.MatchedWithList Then xDriver.BoundText = ""
End Sub
Private Sub xCar_GotFocus()
myGotFocus xCar
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT oil_orders.CODE, CONVERT(VARCHAR(10),oil_orders.DATE,111), CARS.DESCA,TYPE_GAS_CODES.DESCA, DRIVER.DESCA" & _
                  " FROM oil_orders LEFT OUTER JOIN CARS ON oil_orders.CAR = CARS.CODE LEFT OUTER JOIN " & _
                  " DRIVER ON oil_orders.DRIVER = DRIVER.CODE" & _
                  " LEFT JOIN TYPE_GAS_CODES ON oil_orders.GAS = TYPE_GAS_CODES.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by oil_orders.DATE,oil_orders.CODE"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«· «—ÌŒ"
listarray(0, 1) = "(##ORDERS.DATE##)"

listarray(1, 0) = "«·”Ì«—…"
listarray(1, 1) = "%%CARS.DESCA%%"

listarray(2, 0) = "‰Ê⁄ «·ÊﬁÊœ"
listarray(2, 1) = "%%TYPE_GAS_CODES.DESCA%%"

listarray(3, 0) = "«·”«∆ﬁ"
listarray(3, 1) = "%%DRIVER.DESCA%%"

GrdArray(0, 0) = "«·„”·”·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·”Ì«—…"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "‰Ê⁄ «·ÊﬁÊœ"
GrdArray(3, 1) = 1500
GrdArray(4, 0) = "«·”«∆ﬁ"
GrdArray(4, 1) = 2000


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "ORDRES.CODE")
    oSearch.aFilter = aFilter
End If
oSearch.Caption = "≈” ⁄·«„ «Ê«„— «· ‘€Ì·"
oSearch.Show 1
End Sub
Private Sub xGas_LostFocus()
Dim aKind As Variant
aKind = GetFields("select price,kind from type_gas_codes where code = " & xGas.BoundText)
If xGas.MatchedWithList And xCar.MatchedWithList Then
    xDistance_car.Text = Int(Val(xQuant.Text) / Val(GetField("select gas2 from cars where code = " & MyParn(xCar.BoundText)) & ""))
Else
    xDistance.Caption = ""
End If
Calctotals
End Sub
Private Function retPrevious() As Variant
Dim aRet As Variant, cString As String
cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders"
cString = cString & turn(cString) & "CAR = " & MyParn(xCar.BoundText)
cString = cString & turn(cString) & "DATE <= " & DateSq(xDate.Text)
cString = cString & turn(cString) & "COUNTER < " & Val(xCounter.Text)
If Xcode.Text <> "" Then cString = cString & turn(cString) & "oil_orders.CODE <> " & Xcode.Text
cString = cString & " order by date desc,counter desc"
retPrevious = GetFields(cString, con)
End Function
Private Function retNext() As Variant
Dim aRet As Variant, cString As String
cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders inner join type_gas_codes on oil_orders.gas = type_gas_codes.code"
cString = cString & turn(cString) & "car = " & MyParn(xCar.BoundText)
cString = cString & turn(cString) & "DATE >= " & DateSq(xDate.Text)
cString = cString & turn(cString) & "COUNTER > " & Val(xCounter.Text)
If Xcode.Text <> "" Then cString = cString & turn(cString) & "oil_orders.CODE <> " & Xcode.Text
cString = cString & " order by date,counter"
retNext = GetFields(cString, con)
End Function

Private Sub xNote_GotFocus()
myGotFocus xNote
End Sub
Private Sub xNote_KeyUp(KeyCode As Integer, Shift As Integer)
'notesLookupall Me, oSearchNote
End Sub

Private Sub xNote_LostFocus()
myLostFocus xNote
If Not xNote.MatchedWithList Then xNote.BoundText = ""
'If Not xNote.MatchedWithList Then
'    xNote.BoundText = ""
'    XBON.Text = ""
'Else
'    Dim sText As String, bAddText As Boolean
'    sText = XBON.Text
'    XBON.Clear
RetNotes
End Sub
Private Sub xPrice_Change()
Calctotals
End Sub

Private Sub xQuant_Change()
Calctotals
End Sub
Private Function FixOrder() As Boolean
Dim aRet As Variant, cString As String
Dim aInsert As Variant
If Xcode.Tag = LoadMode Then
    aRow = GetFields("select * from oil_orders where code = " & MyParn(Xcode.Text), con)
    If IsEmpty(aRow) Then Exit Function
    
    cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders"
    cString = cString & turn(cString) & "car = " & MyParn(retFlag(aRow, "car"))
    cString = cString & turn(cString) & "DATE <= " & DateSq(retFlag(aRow, "DATE"))
    cString = cString & turn(cString) & "COUNTER < " & DateSq(retFlag(aRow, "COUNTER"))
    cString = cString & " order by date desc,counter desc"
    aPrevious = GetFields(cString)
    
    cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders"
    cString = cString & turn(cString) & "car = " & MyParn(retFlag(aRow, "car"))
    cString = cString & turn(cString) & "DATE >= " & DateSq(retFlag(aRow, "DATE"))
    cString = cString & turn(cString) & "COUNTER > " & DateSq(retFlag(aRow, "COUNTER"))
    cString = cString & " order by date,counter"
    aNext = GetFields(cString)

    con.BeginTrans
    On Error GoTo myerror
    If Not IsEmpty(aPrevious) Then
        If Not IsEmpty(aNext) Then
            aInsert = AddFlag(Empty, "code_next", addstring(retFlag(aNext, "code")))
            aInsert = AddFlag(aInsert, "Counter_Next", Val(retFlag(aNext, "counter")))
        Else
            aInsert = AddFlag(Empty, "code_next", "null")
            aInsert = AddFlag(aInsert, "Counter_Next", "0")
        End If
        con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aPrevious, "code")))
    End If

    If Not IsEmpty(aNext) Then
        If Not IsEmpty(aPrevious) Then
            aInsert = AddFlag(Empty, "code_Prv", addstring(retFlag(aPrevious, "code")))
        Else
            aInsert = AddFlag(Empty, "code_Prv", "null")
        End If
        con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aNext, "code")))
    End If
End If

aPrevious = retPrevious
aNext = retNext

aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "Quant", Val(xQuant.Text))
aInsert = AddFlag(aInsert, "Quant2", Val(xQuant2.Text))
aInsert = AddFlag(aInsert, "bon", addstring(xBon.Text))
aInsert = AddFlag(aInsert, "Driver", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "car_Driver", addstring(xcar_driver.BoundText))
aInsert = AddFlag(aInsert, "Gas", addvalue(xGas.BoundText))
aInsert = AddFlag(aInsert, "Gas2", addvalue(xGas2.BoundText))
aInsert = AddFlag(aInsert, "TOTAL", Val(xTotal.Text))
aInsert = AddFlag(aInsert, "car", addvalue(xCar.BoundText))
aInsert = AddFlag(aInsert, "Counter", Val(xCounter.Text))
aInsert = AddFlag(aInsert, "Distance_car", Val(xDistance_car.Text))
aInsert = AddFlag(aInsert, "code_prv", addstring(retFlag(aPrevious, "code")))
If Not IsEmpty(aNext) Then
    aInsert = AddFlag(aInsert, "code_Next", addstring(retFlag(aNext, "code")))
    aInsert = AddFlag(aInsert, "counter_Next", Val(retFlag(aNext, "counter")))
Else
    aInsert = AddFlag(aInsert, "code_Next", "Null")
    aInsert = AddFlag(aInsert, "counter_Next", "0")
End If

If Xcode.Tag = LoadMode Then
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(Xcode.Text))
Else
    con.BeginTrans
    On Error GoTo myerror
    Xcode.Text = RetZero(Val(Newflag("oil_orders", "code")))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(Xcode.Text))
    con.Execute addInsert(aInsert, "oil_orders")
End If

If Not IsEmpty(aPrevious) Then
    aInsert = AddFlag(Empty, "code_next", addstring(Xcode.Text))
    aInsert = AddFlag(aInsert, "Counter_Next", Val(xCounter.Text))
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aPrevious, "code")))
End If
'
If Not IsEmpty(aNext) Then
    aInsert = AddFlag(Empty, "code_prv", addstring(Xcode.Text))
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aNext, "code")))
End If
con.CommitTrans
FixOrder = True
Exit Function
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Function FixOrderDelete() As Boolean
Dim aRet As Variant, cString As String
Dim aInsert As Variant
aRow = GetFields("select * from oil_orders where code = " & MyParn(Xcode.Text), con)
If IsEmpty(aRow) Then Exit Function


cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders"
cString = cString & turn(cString) & "car = " & MyParn(retFlag(aRow, "car"))
cString = cString & turn(cString) & "DATE <= " & DateSq(retFlag(aRow, "DATE"))
cString = cString & turn(cString) & "COUNTER < " & DateSq(retFlag(aRow, "COUNTER"))
cString = cString & " order by date desc,counter desc"
aPrevious = GetFields(cString)

cString = "select TOP 1 oil_orders.CODE, oil_orders.counter,Distance from oil_orders"
cString = cString & turn(cString) & "car = " & MyParn(retFlag(aRow, "car"))
cString = cString & turn(cString) & "DATE >= " & DateSq(retFlag(aRow, "DATE"))
cString = cString & turn(cString) & "COUNTER > " & DateSq(retFlag(aRow, "COUNTER"))
cString = cString & " order by date,counter"
aNext = GetFields(cString)

con.BeginTrans
If Not IsEmpty(aPrevious) Then
    If Not IsEmpty(aNext) Then
        aInsert = AddFlag(Empty, "code_next", addstring(retFlag(aNext, "code")))
        aInsert = AddFlag(aInsert, "Counter_Next", Val(retFlag(aNext, "counter")))
    Else
        aInsert = AddFlag(Empty, "code_next", "null")
        aInsert = AddFlag(aInsert, "Counter_Next", "0")
    End If
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aPrevious, "code")))
End If

If Not IsEmpty(aNext) Then
    If Not IsEmpty(aPrevious) Then
        aInsert = AddFlag(Empty, "code_Prv", addstring(retFlag(aPrevious, "code")))
    Else
        aInsert = AddFlag(Empty, "code_Prv", "null")
    End If
    con.Execute addUpdate(aInsert, "oil_orders", "CODE = " & addstring(retFlag(aNext, "code")))
End If

con.Execute "Delete  From oil_orders  Where code = " & MyParn(Xcode.Text)
con.CommitTrans
FixOrderDelete = True
Exit Function
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Function validPrevious() As String
Dim aRet As Variant, cString As String
cString = "select top 1 counter,oil_orders.code from oil_orders inner join type_gas_codes on oil_orders.gas = type_gas_codes.code"
cString = cString & turn(cString) & "car = " & MyParn(xCar.BoundText)
cString = cString & turn(cString) & "DATE < " & DateSq(xDate.Text)
cString = cString & turn(cString) & "COUNTER >= " & Val(xCounter.Text)
If Xcode.Text <> "" Then cString = cString & turn(cString) & "oil_orders.CODE <> " & Xcode.Text

aRet = GetFields(cString, con)
If Not IsEmpty(aRet) Then
    validPrevious = retFlag(aRet, "code")
Else
    validPrevious = "ok"
End If
End Function
Private Function validNext() As String
Dim aRet As Variant, cString As String
cString = "select  top 1 Counter,oil_orders.code from oil_orders inner join type_gas_codes on oil_orders.gas = type_gas_codes.code"
cString = cString & turn(cString) & "car = " & MyParn(xCar.BoundText)
cString = cString & turn(cString) & "DATE > " & DateSq(xDate.Text)
cString = cString & turn(cString) & "COUNTER <= " & Val(xCounter.Text)
If Xcode.Text <> "" Then cString = cString & turn(cString) & "oil_orders.CODE <> " & Xcode.Text
aRet = GetFields(cString, con)
If Not IsEmpty(aRet) Then
    validNext = retFlag(aRet, "code")
Else
    validNext = "ok"
End If
End Function
Private Sub RetNotes()
If Not xNote.MatchedWithList Then
    xBons.Caption = ""
    xBons_used.Caption = ""
    xBons_rest.Caption = ""
    Exit Sub
End If
Dim aRet As Variant, cString As String
aRet = GetFields("select BON,BON_COUNT FROM notes_codes_oil WHERE CODE = " & xNote.BoundText)
If Not IsEmpty(aRet) Then
    xBons.Caption = Myvalue(retFlag(aRet, "bon_count"))
    xBons_used.Caption = Myvalue(GetField("select COUNT(*) FROM oil_orders WHERE NOTE = " & xNote.BoundText))
    xBons_rest.Caption = Myvalue(Val(xBons.Caption) - Val(xBons_used.Caption))
End If
End Sub
Private Sub xQuant_LostFocus()
'If xGas.MatchedWithList And xcar.MatchedWithList Then
'    xDistance_car.Text = Int(Val(xQuant.Text) / Val(GetField("select gas2 from cars where code = " & MyParn(xcar.BoundText)) & ""))
'Else
'    xDistance_car.Text = ""
'End If
End Sub

