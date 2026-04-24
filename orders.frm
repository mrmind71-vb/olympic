VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ordersfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«Ê«„— «·‘€·"
   ClientHeight    =   8595
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   11085
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
   ScaleHeight     =   8595
   ScaleWidth      =   11085
   Begin VB.Frame Frame6 
      Height          =   1005
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   4095
      Width           =   10815
      Begin VB.TextBox xDistance_v 
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
         Left            =   6075
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Tag             =   "aa"
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "›—ﬁ »‰“Ì‰ „› —÷"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   585
         Width           =   1470
      End
      Begin VB.Label xGas_differ 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "›—ﬁ «·„”«›…"
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
         TabIndex        =   58
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label xDistance_differ 
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
         TabIndex        =   57
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "«” Â·«ﬂ »‰“Ì‰ „› —÷"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   225
         Width           =   1785
      End
      Begin VB.Label xGas_v 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "„”«›… „› —÷…"
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
         TabIndex        =   54
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1680
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   2430
      Width           =   10815
      Begin VB.TextBox xCounter_out 
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
         Left            =   6030
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "aa"
         Top             =   180
         Width           =   2940
      End
      Begin VB.TextBox xTime_out 
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
         Left            =   135
         MaxLength       =   5
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "aa"
         Top             =   180
         Width           =   2940
      End
      Begin VB.TextBox xCounter_in 
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
         Left            =   6030
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "aa"
         Top             =   540
         Width           =   2940
      End
      Begin VB.TextBox xTime_in 
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
         Left            =   135
         MaxLength       =   5
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Tag             =   "aa"
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "«” Â·«ﬂ »‰“Ì‰ ·ﬂ· ﬂÌ·Ê „ —"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1305
         Width           =   2145
      End
      Begin VB.Label xgas1 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1260
         Width           =   2940
      End
      Begin VB.Label xGas 
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1260
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "«” Â·«ﬂ »‰“Ì‰  ﬁœÌ—Ì"
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
         TabIndex        =   50
         Top             =   1305
         Width           =   1710
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ ⁄œ«œ «·Œ—ÊÕ"
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
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄… «·Œ—ÊÃ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ ⁄œ«œ «·⁄Êœ…"
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
         TabIndex        =   47
         Top             =   585
         Width           =   1200
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "”«⁄… «·⁄Êœ…"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·„”«›… »«·ﬂÌ·Ê"
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
         TabIndex        =   45
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label xPeriod 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·„œ…"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   945
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2670
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   5130
      Width           =   10770
      Begin VB.TextBox xLine 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   135
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "new"
         Top             =   450
         Width           =   10500
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   45
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Tag             =   "new"
         Top             =   1710
         Width           =   10590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Œÿ «·”Ì—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  «·Ã—«Ã"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1395
         Width           =   1290
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
      TabIndex        =   31
      Top             =   9180
      Visible         =   0   'False
      Width           =   7800
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3300
         Left            =   135
         TabIndex        =   15
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
      Height          =   1725
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   720
      Width           =   10770
      Begin VB.TextBox xOrder_No 
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
         Left            =   5985
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1305
         Width           =   2940
      End
      Begin VB.CheckBox xMission 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„√„Ê—Ì…"
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   990
         Width           =   1050
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   225
         Width           =   2940
      End
      Begin VB.TextBox xEmp 
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
         Left            =   90
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1305
         Width           =   4380
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
         Left            =   7830
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo xcar_driver 
         Height          =   330
         Left            =   90
         TabIndex        =   3
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
         Left            =   5985
         TabIndex        =   4
         Top             =   945
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
         Left            =   5985
         TabIndex        =   2
         Top             =   585
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «„— «·‘€·"
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
         TabIndex        =   63
         Top             =   1350
         Width           =   1125
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
         TabIndex        =   41
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "«·„—«›ﬁÌ‰"
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1350
         Width           =   735
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
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   945
         Width           =   1335
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
         TabIndex        =   35
         Top             =   1710
         Width           =   60
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   34
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   570
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
         Left            =   9045
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7740
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   24
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
         Picture         =   "orders.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "orders.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   810
         TabIndex        =   25
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
         Picture         =   "orders.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "orders.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
         TabIndex        =   26
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
         Picture         =   "orders.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "orders.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2340
         TabIndex        =   27
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
         Picture         =   "orders.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "orders.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   17
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
         Picture         =   "orders.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "orders.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "orders.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "orders.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "orders.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "orders.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   1080
      Top             =   990
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
      Left            =   360
      Top             =   90
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
      Left            =   405
      Top             =   360
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
      BackColor       =   32768
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
      Left            =   1980
      TabIndex        =   32
      Top             =   7830
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
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7830
      Width           =   4425
   End
End
Attribute VB_Name = "Ordersfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bedit As Boolean
Dim aEditRecord As Variant
Public sCode As String
Dim nRound As Long
Dim con As New ADODB.Connection
Dim cFilter As String, cFilterLookup As String
Dim oSearch As New Search3, oSearchCode As New Search3, oSearchCharge As New Search3
Dim formMode As Byte, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdFilter_Click()
cFilterLookup = ""
openCardTable
myUndo
End Sub
Private Sub cmdGroup_Click()
Dim oFlagGroup As New FlagGroupFrm, sCode As String
sCode = xcar.BoundText
oFlagGroup.sCaption = "„Ã„Ê⁄«  «·«’‰«›"
oFlagGroup.sCode = "«·ﬂÊœ"
oFlagGroup.sDesca = "≈”„ «·„Ã„Ê⁄…"
oFlagGroup.sGroupDesca = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
oFlagGroup.sTable = "FILE1_50"
oFlagGroup.sTableGroup = "FILE1_50G"
oFlagGroup.nZero = -1
oFlagGroup.nZeroGroup = -1
oFlagGroup.sGroupCaption = "„Ã„Ê⁄«  «·«’‰«› «·—∆Ì”Ì…"
oFlagGroup.Show 1
DATA1.Refresh
If sCode <> "" Then xcar.BoundText = sCode
If Not xcar.MatchedWithList Then xcar.BoundText = ""
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Activate()
If sCode <> "" Then
    On Error Resume Next
    xCode.SetFocus
    Err.Clear
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
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con

Set DATA1.Recordset = myRecordSet("SELECT * FROM CARS", con)
Set xcar.RowSource = DATA1
xcar.ListField = "Desca"
xcar.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("SELECT * FROM DRIVER", con)
Set xDriver.RowSource = DATA2
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

Set xcar_driver.RowSource = DATA2
xcar_driver.ListField = "Desca"
xcar_driver.BoundColumn = "Code"


openCardTable
myUndo
If sCode = "" And aEditRecord Then CmdAdd_Click
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.Text = ""
On Error Resume Next
xDate.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From ORDERS  Where CODE = " & MyParn(xCode.Text)
    con.CommitTrans
    If sCode <> "" Then
        Unload Me
        Exit Sub
    End If
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "CODE < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
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
    con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If sCode <> "" Then
    Unload Me
    Exit Sub
End If
If xCode.Tag = DefineMode Then
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
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = ""
xDate.Text = ""
xcar.BoundText = "1"
xcar_driver.BoundText = ""
xOrder_No.Text = Val(xOrder_No.Text) + 1
xCounter_out.Text = ""
xTime_out.Text = ""
xCounter_in.Text = ""
xTime_in.Text = ""
xDriver.BoundText = ""
xEmp.Text = ""
xLine.Text = ""
xNotes.Text = ""
xDistance.Caption = ""
xDistance_v.Text = ""
xPeriod.Caption = ""
xPeriod.Tag = ""
xDistance.Caption = ""
xGas.Caption = ""
xGas_v.Caption = ""
xGas_differ.Caption = ""
xDistance_differ.Caption = ""
xgas1.Caption = ""
xMission.Value = 0
Handlecontrols DefineMode
xRecord.Caption = "«÷«›… ”Ã· ÃœÌœ"
End Sub
Private Sub myload()
xCode.Text = CardTable!code & ""
xDate.Text = Format(CardTable!Date, "DD-MM-YYYY")
xcar.BoundText = CardTable!car & ""
xcar_driver.BoundText = CardTable!car_driver & ""
xCounter_out.Text = CardTable!Counter_out & ""
xOrder_No.Text = CardTable!Order_No & ""
xTime_out.Text = RetTime(CardTable!Time_out & "")
xCounter_in.Text = CardTable!Counter_in & ""
xTime_in.Text = RetTime(CardTable!Time_in & "")
xDriver.BoundText = CardTable!Driver & ""
xEmp.Text = CardTable!Emp & ""
xLine.Text = CardTable!Line & ""
xNotes.Text = CardTable!Notes & ""
xDistance.Caption = CardTable!Distance & ""
xDistance_v.Text = CardTable!Distance_v & ""
xgas1.Caption = CardTable!gas1
xMission.Value = IIf(CardTable!MISSION, 1, 0)
Calctotals
Calctotals2
xRecord.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
Handlecontrols LoadMode
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Calctotals
Calctotals2
Dim aInsert As Variant, bAddBar As Boolean
aInsert = AddFlag(Empty, "[DATE]", addDate(xDate.Text))
aInsert = AddFlag(aInsert, "car", addvalue(xcar.BoundText))
aInsert = AddFlag(aInsert, "Order_no", addvalue(xOrder_No.Text))
aInsert = AddFlag(aInsert, "car_driver", addstring(xcar_driver.BoundText))
aInsert = AddFlag(aInsert, "Counter_out", addstring(xCounter_out.Text))
aInsert = AddFlag(aInsert, "Time_out", addstring(xTime_out.Text))
aInsert = AddFlag(aInsert, "Counter_in", addstring(xCounter_in.Text))
aInsert = AddFlag(aInsert, "Time_in", addstring(xTime_in.Text))
aInsert = AddFlag(aInsert, "Driver", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "Distance", Val(xDistance.Caption))
aInsert = AddFlag(aInsert, "Period", Val(xPeriod.Caption))
aInsert = AddFlag(aInsert, "Emp", addstring(xEmp.Text))
aInsert = AddFlag(aInsert, "Line", addstring(xLine.Text))
aInsert = AddFlag(aInsert, "Notes", addstring(xNotes.Text))
aInsert = AddFlag(aInsert, "Distance_v", Val(xDistance_v.Text))
aInsert = AddFlag(aInsert, "gas1", Val(xgas1.Caption))
aInsert = AddFlag(aInsert, "Mission", xMission.Value)
con.BeginTrans
If xCode.Text = "" Then
    xCode.Text = RetZero(Val(Newflag("ORDERS", "code")))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(xCode.Text))
    con.Execute addInsert(aInsert, "ORDERS")
Else
    con.Execute addUpdate(aInsert, "ORDERS", "CODE = " & addstring(xCode.Text))
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
    xCode.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)
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

Private Sub xcar_Click(Area As Integer)
If Area = 2 Then xCar_LostFocus
End Sub

Private Sub xCar_LostFocus()
myLostFocus xcar
If xcar.MatchedWithList Then
    Dim aRet As Variant
    aRet = GetFields("select driver,line,distance,gas1,quant from cars where code = " & xcar.BoundText, con)
    xcar_driver.BoundText = retFlag(aRet, "driver") & ""
    xDriver.BoundText = retFlag(aRet, "driver") & ""
    xLine.Text = retFlag(aRet, "line") & ""
    xgas1.Caption = Myvalue(retFlag(aRet, "gas1") & "")
    xDistance_v.Text = Myvalue(retFlag(aRet, "Distance") & "")
Else
    xcar.BoundText = ""
    xDriver.BoundText = ""
    xcar_driver.BoundText = ""
    xLine.Text = ""
    xgas1.Caption = ""
    xDistance_v.Text = ""
End If
Calctotals
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text)
CardTable.Find "code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    If xCode.Tag = LoadMode Then mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not xcar.MatchedWithList Then
    If Not bIgMsg Then MsgBox "«·”«∆ﬁ €Ì— „”Ã·"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xCode.Text) <> "" Then
        CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT * FROM ORDERS"
cFilter = ""
If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "ORDERS.CODE = " & MyParn(sCode)
If cFilterLookup <> "" Then cFilter = cFilter & turn(cFilter, " and ") & cFilterLookup
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY ORDERS.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub Calctotals()
xDistance.Caption = Myvalue(Val(xCounter_in.Text) - Val(xCounter_out.Text))
xGas.Caption = Myvalue(Val(xDistance.Caption) * Val(xgas1.Caption))
xGas_v.Caption = Myvalue(Val(xDistance_v.Text) * Val(xgas1.Caption))
xGas_differ.Caption = Myvalue(Val(xGas.Caption) - Val(xGas_v.Caption))
xDistance_differ.Caption = Myvalue(Val(xDistance.Caption) - Val(xDistance_v.Text))
If Val(xGas_differ.Caption) > 0 Then xGas_differ.Caption = "+" & xGas_differ.Caption
If Val(xDistance_differ.Caption) > 0 Then xDistance_differ.Caption = "+" & xDistance_differ.Caption
End Sub
Private Sub Calctotals2()
xTime_in.Text = RetTime(xTime_in.Text)
xTime_out.Text = RetTime(xTime_out.Text)
If IsTime(xTime_in.Text) And IsTime(xTime_out.Text) Then
    xPeriod.Caption = MinuteToTimeDg(retDiffMinutes(xTime_out.Text, xTime_in.Text, True))
    xPeriod.Tag = retDiffMinutes(xTime_out.Text, xTime_in.Text, True)
Else
    xPeriod.Caption = ""
    xPeriod.Tag = ""
End If
End Sub
Private Sub xCounter_in_LostFocus()
myLostFocus xCounter_in
Calctotals
End Sub
Private Sub xCounter_out_LostFocus()
myLostFocus xCounter_out
Calctotals
End Sub

Private Sub xDistance_differ_Change()
xDistance_differ.ForeColor = IIf(Val(xDistance_differ.Caption) > 0, vbRed, vbBlack)
xGas_differ.ForeColor = IIf(Val(xGas_differ.Caption) > 0, vbRed, vbBlack)
End Sub

Private Sub xDistance_v_Change()
Calctotals
End Sub

Private Sub xTime_in_LostFocus()
myLostFocus xTime_in
Calctotals2
End Sub
Private Sub xTime_out_LostFocus()
myLostFocus xTime_out
Calctotals2
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
myGotFocus xCode
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
myGotFocus xcar
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT ORDERS.CODE, CONVERT(VARCHAR(10),ORDERS.DATE,111), CARS.DESCA, DRIVER.DESCA, ORDERS.LINE" & _
                  " FROM ORDERS LEFT OUTER JOIN CARS ON ORDERS.CAR = CARS.CODE LEFT OUTER JOIN " & _
                  " DRIVER ON ORDERS.DRIVER = DRIVER.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by ORDERS.DATE,ORDERS.CODE"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«· «—ÌŒ"
listarray(0, 1) = "(##ORDERS.DATE##)"

listarray(1, 0) = "«·”Ì«—…"
listarray(1, 1) = "%%CARS.DESCA%%"

listarray(2, 0) = "«·”«∆ﬁ"
listarray(2, 1) = "%%DRIVER.DESCA%%"

listarray(3, 0) = "Œÿ «·”Ì—"
listarray(3, 1) = "%%ORDERS.LINE%%"

GrdArray(0, 0) = "«·„”·”·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·”Ì«—…"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·”«∆ﬁ"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "Œÿ «·”Ì—"
GrdArray(4, 1) = 8000

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
Private Sub xDistance_v_GotFocus()
myGotFocus xDistance_v
End Sub
Private Sub xDistance_v_LostFocus()
myLostFocus xDistance_v
End Sub

