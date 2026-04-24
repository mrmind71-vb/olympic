VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form memberfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "»Ì«‰«  «⁄÷«¡ «·‰«œÌ"
   ClientHeight    =   9720
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
   ScaleHeight     =   9720
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   124
      Top             =   0
      Width           =   1950
      Begin Threed.SSCommand cmdDocument 
         Height          =   510
         Left            =   45
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
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
         Picture         =   "Member.frx":0000
         Caption         =   "Õ›Ÿ «·„” ‰œ« "
         Alignment       =   1
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member.frx":264E
      End
   End
   Begin VB.CheckBox chkHideInstall 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "«Œ›«¡ «·„”Ã·"
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
      Left            =   15930
      RightToLeft     =   -1  'True
      TabIndex        =   120
      Top             =   8370
      Width           =   1680
   End
   Begin VB.CheckBox chkInstall 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "√⁄÷«¡ ⁄·ÌÂ„ ﬁÌ„… „÷«›…"
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
      Left            =   17820
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   8370
      Width           =   2310
   End
   Begin VB.CommandButton Command511 
      Caption         =   "Command5"
      Height          =   375
      Left            =   -1530
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   2250
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Timer CardTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10935
      Top             =   8775
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   450
      Width           =   4875
      Begin Threed.SSCommand cmdClaimFawry 
         Height          =   825
         Left            =   90
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   585
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1455
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
         Picture         =   "Member.frx":4C9C
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":812D
      End
      Begin Threed.SSCommand cmdClaim 
         Height          =   870
         Left            =   90
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1440
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1535
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
         Picture         =   "Member.frx":AE88
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":D349
      End
      Begin VB.Label xPaid_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   1980
         Width           =   2040
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·”œ«œ"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   2025
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  «—ÌŒ «Œ— ”œ«œ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label xdate_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «‰ Â«¡ «· ÃœÌœ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   945
         Width           =   1530
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «Ì’«· «·œ›⁄"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   1665
         Width           =   1500
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   1620
         Width           =   2040
      End
      Begin VB.Label xDate_End 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   900
         Width           =   2040
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… «Ì’«· «·œ›⁄"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   1305
         Width           =   1635
      End
      Begin VB.Label xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   1260
         Width           =   2040
      End
      Begin VB.Label xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   540
         Width           =   2040
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ »œ«Ì… «· ÃœÌœ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   585
         Width           =   1530
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   89
      Top             =   2790
      Width           =   4875
      Begin VB.CheckBox xInstall 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "⁄÷Ê  ﬁ”Ìÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   585
         Width           =   1410
      End
      Begin VB.CheckBox xapg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌÀ »ÿ«ﬁ« "
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   225
         Width           =   1410
      End
      Begin VB.CheckBox xDied 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "„ Ê›Ì"
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   225
         Width           =   870
      End
      Begin VB.CheckBox xDrop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "”«ﬁÿ ⁄÷ÊÌ…"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   3780
      Width           =   9420
      Begin VB.CommandButton cmdCardReader 
         Caption         =   "..."
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ﬂ«—‰ÌÂ"
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
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label xCard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   495
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ ÿ»«⁄… «·ﬂ«—‰ÌÂ"
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
         Left            =   7650
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label xDate_print 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   270
         Width           =   2130
      End
   End
   Begin VB.Frame Frame11 
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
      Height          =   1050
      Left            =   9495
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   4545
      Width           =   10725
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
         Height          =   720
         Left            =   90
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   9195
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
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   4590
      Width           =   9375
      Begin VB.TextBox xCode_main 
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
         TabIndex        =   25
         Tag             =   "D"
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox xDate_Begin 
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
         Left            =   5805
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "D"
         Top             =   540
         Width           =   2130
      End
      Begin VB.TextBox xSes_no 
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
         TabIndex        =   23
         Tag             =   "D"
         Top             =   180
         Width           =   2130
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   5805
         TabIndex        =   22
         Top             =   180
         Width           =   2130
         _ExtentX        =   3757
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
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›«’· „‰"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„Ê«›ﬁ…"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "›∆… «·⁄÷ÊÌ…"
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
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ «·„Ê«›ﬁ…"
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
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   585
         Width           =   1275
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
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   11565
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   8655
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   3780
         TabIndex        =   64
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
         Picture         =   "Member.frx":100A4
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":12A99
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   65
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
         Picture         =   "Member.frx":15332
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   1395
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "Member.frx":17655
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":19DF1
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   2430
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
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
         Picture         =   "Member.frx":1C285
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":1E4C6
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   510
         Left            =   4950
         TabIndex        =   68
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
         Picture         =   "Member.frx":207B3
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":227BB
      End
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   7515
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "Member.frx":24772
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":26B3D
      End
      Begin Threed.SSCommand cmdInform_rel 
         Height          =   510
         Left            =   6210
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
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
         Caption         =   "«” ⁄·«„  «»⁄"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":28BE6
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
      Height          =   2850
      Left            =   9540
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   720
      Width           =   10680
      Begin VB.CheckBox xChair 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄÷Ê „Ã·” «œ«—…"
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
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   990
         Width           =   1815
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
         Left            =   495
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2385
         Width           =   3615
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
         Left            =   5490
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2385
         Width           =   3750
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
         Height          =   330
         Left            =   495
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2025
         Width           =   8745
      End
      Begin VB.CommandButton cmdRegion 
         Caption         =   "..."
         Height          =   330
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1305
         Width           =   375
      End
      Begin VB.CommandButton cmdDegree 
         Caption         =   "..."
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1665
         Width           =   375
      End
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
         Left            =   5895
         MaxLength       =   14
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1665
         Width           =   3345
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
         Left            =   7380
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   945
         Width           =   1860
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
         Left            =   540
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2310
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
         Left            =   5805
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   3435
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
         Left            =   7695
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "2"
         Top             =   225
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   540
         TabIndex        =   9
         Top             =   1305
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   540
         TabIndex        =   6
         Top             =   945
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   540
         TabIndex        =   3
         Top             =   585
         Width           =   2310
         _ExtentX        =   4075
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
         Left            =   540
         TabIndex        =   11
         Top             =   1665
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSDataListLib.DataCombo xRegion 
         Height          =   330
         Left            =   6255
         TabIndex        =   7
         Top             =   1305
         Width           =   2985
         _ExtentX        =   5265
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
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   " ·Ì›Ê‰ «·„‰“·"
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   2070
         Width           =   1035
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Õ· «·«ﬁ«„…"
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1368
         Width           =   1080
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1719
         Width           =   1035
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
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
         Height          =   240
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   46
         Top             =   1305
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   270
         Width           =   585
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   270
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   621
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
      Left            =   10800
      Top             =   6705
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   375
      Left            =   -1305
      Top             =   1395
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
      Left            =   -1215
      Top             =   1170
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   375
      Left            =   -1440
      Top             =   1485
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
      Left            =   -1260
      Top             =   1260
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
      Left            =   -1575
      Top             =   1305
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
      Left            =   -1305
      Top             =   1125
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   375
      Left            =   -1575
      Top             =   1035
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   3525
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   225
      Width           =   4470
      Begin Threed.SSCommand cmdScan 
         Height          =   555
         Left            =   90
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4290
         _ExtentX        =   7567
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
         Picture         =   "Member.frx":2AB9D
         Caption         =   "„”Õ ÷Ê∆Ì"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "Member.frx":2D337
      End
      Begin VB.Image xAppendPhoto 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2670
         Left            =   90
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2130
      End
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2670
         Left            =   2250
         Stretch         =   -1  'True
         Top             =   180
         Width           =   2130
      End
   End
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "Member.frx":30092
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   10755
      TabIndex        =   37
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   9540
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3555
      Width           =   10680
      Begin VB.TextBox xPhone_work 
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
         Left            =   5535
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   3705
      End
      Begin VB.CommandButton cmdCompany 
         Caption         =   "..."
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox xJob_desca 
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
         TabIndex        =   20
         Tag             =   "D"
         Top             =   585
         Width           =   3750
      End
      Begin VB.CommandButton cmdJob 
         Caption         =   "..."
         Height          =   330
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin MSDataListLib.DataCombo xJob 
         Height          =   330
         Left            =   5895
         TabIndex        =   16
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
      Begin MSDataListLib.DataCombo xCompany 
         Height          =   330
         Left            =   495
         TabIndex        =   18
         Top             =   225
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
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   " ·Ì›Ê‰ «·⁄„·"
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·‘—ﬂ…"
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄‰Ê«‰ «·⁄„·"
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9315
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   180
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc data9 
      Height          =   420
      Left            =   4455
      Top             =   -90
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   135
      TabIndex        =   30
      Top             =   5625
      Width           =   20115
      _ExtentX        =   35481
      _ExtentY        =   5265
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " ”·Ì„ «·ﬂ«—‰ÌÂ« "
      TabPicture(0)   =   "Member.frx":304D4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdDeliverAll"
      Tab(0).Control(1)=   "grid2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "„ÿ«·»«  «·⁄÷Ê"
      TabPicture(1)   =   "Member.frx":304F0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "»Ì«‰«  «· Ê«»⁄"
      TabPicture(2)   =   "Member.frx":3050C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid1"
      Tab(2).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   2535
         Left            =   -70995
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   360
         Width           =   16035
         _cx             =   28284
         _cy             =   4471
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   2535
         Left            =   -74910
         TabIndex        =   26
         Top             =   360
         Width           =   19950
         _cx             =   35190
         _cy             =   4471
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
      Begin VSFlex7Ctl.VSFlexGrid grid3 
         Height          =   2535
         Left            =   90
         TabIndex        =   27
         Top             =   360
         Width           =   19905
         _cx             =   35110
         _cy             =   4471
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
      Begin Threed.SSCommand cmdDeliverAll 
         Height          =   420
         Left            =   -74910
         TabIndex        =   126
         Top             =   2475
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
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
         Caption         =   " ”·Ì„ «·ﬂ·"
         TagVariant      =   "«Œ «— « ·Œ«„…"
         ButtonStyle     =   3
      End
   End
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   375
      Left            =   3780
      Top             =   -180
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   75
      Top             =   9255
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
         TabIndex        =   76
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
         TabIndex        =   77
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
         TabIndex        =   78
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
         TabIndex        =   79
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
         TabIndex        =   80
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   8595
      Width           =   3660
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2745
         TabIndex        =   71
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
         Picture         =   "Member.frx":30528
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member.frx":326CF
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1845
         TabIndex        =   72
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
         Picture         =   "Member.frx":34716
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member.frx":36801
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   990
         TabIndex        =   73
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
         Picture         =   "Member.frx":387FB
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member.frx":3A90C
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   90
         TabIndex        =   74
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
         Picture         =   "Member.frx":3C906
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "Member.frx":3EB2A
      End
   End
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   420
      Left            =   -1215
      Top             =   2025
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
   Begin Olymbic.CSubclass CSubclass1 
      Left            =   -360
      Top             =   810
      _ExtentX        =   1693
      _ExtentY        =   1535
   End
   Begin Threed.SSCommand cmdCard 
      Height          =   420
      Left            =   18990
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   9000
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Caption         =   "‘Õ‰ «·ﬂ«—‰ÌÂ"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "Member.frx":40BFB
   End
   Begin Threed.SSCommand cmdSendAll 
      Height          =   420
      Left            =   16200
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   8775
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
      Caption         =   "‰ﬁ· ﬂ· «·»Ì«‰«  ··»Ê«»…"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "Member.frx":42BB2
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   8595
      Width           =   3840
      Begin VB.CheckBox chkFawry 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«⁄÷«¡ ·Â„ Õ”«» ›Ï ›Ê—Ì"
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
         Height          =   330
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   111
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label xFawryValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   8595
      Width           =   1950
      Begin VB.CheckBox chkFawryHidden 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«⁄÷«¡  ”ÊÌ« "
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
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   225
         Width           =   1590
      End
   End
   Begin Threed.SSCommand cmdCardTrans 
      Height          =   420
      Left            =   14175
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
      Caption         =   "«—”«· «·»Ì«‰«  «·Ì «·»Ê«»…"
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "Member.frx":44B69
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   8595
      Width           =   3795
      Begin VB.CheckBox chkCurrent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·Õ«·Ì"
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
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   123
         Top             =   180
         Value           =   1  'Checked
         Width           =   825
      End
      Begin Threed.SSCommand cmdService 
         Height          =   465
         Left            =   1260
         TabIndex        =   122
         Top             =   135
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   820
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
         Caption         =   "⁄—÷"
         TagVariant      =   "«Œ «— « ·Œ«„…"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "memberfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, bEditRecord As Boolean
Public nTab As Integer
Dim con As New ADODB.Connection, aRecords As Variant
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nPhoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String, bAct As Boolean
Dim formMode As Byte
Dim oSearch As New Search, oSearchRel As New Search, oSearchClaim As New Search_empty, oSearchService As New Search_empty
Dim CardTable As ADODB.Recordset
Dim TimerMode As Integer
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Dim bCheck As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bedit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
cmddel.Enabled = (nMode = LoadMode And bEditRecord) And xDrop.Value = 1
cmdSave.Enabled = bEditRecord
cmdInform.Enabled = (nMode = LoadMode)
cmdScan.Enabled = nMode = LoadMode And bEditRecord
cmdService.Enabled = nMode = LoadMode
cmdDocument.Enabled = nMode = LoadMode
aRecords = retRecords(xCode.text)
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
'xCode.Enabled = bEdit And Not (nMode = LoadMode)
'xCode.Enabled = False

'cmdScan2.Enabled = nMode = LoadMode And bEditRecord
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.text = Newflag("file1_10", "code")
xChair.Value = 0
xTitle.text = ""
xDesca.text = ""
xFawryValue.Caption = ""
xCard.Caption = ""
xCode_main.text = ""
xDied.Value = 0

bCheck = False
xInstall.Value = 0
bCheck = True

xJob_desca.text = ""
xDrop.Value = 0
xapg.Value = 0
xNotes.text = ""
xDate_birth.text = ""
xCompany.BoundText = ""
xGender.BoundText = "1"
xSocial.BoundText = ""
'xReason.BoundText = ""
xDate_Begin.text = ""
'xDate_Join.Text = ""
'xDate_Trans.Text = ""
xSes_no.text = ""

'xFace.Text = ""
xReligion.BoundText = "1"
xId_no.text = ""
xAddress.text = ""
xPhone.text = ""
xMobil.text = ""
'xMail.Text = ""
xJob.BoundText = ""
xPhone_work.text = ""
xDegree.BoundText = ""
xRegion.BoundText = ""
xType.BoundText = ""
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")


xDate_print.Caption = ""
xDate_Paid.Caption = ""
xDate1.Caption = ""
xdate_end.Caption = ""
xPaid_desca.Caption = ""

xDoc_No.Caption = ""

panel1(0).Caption = ""
panel1(1).Caption = ""
panel1(2).Caption = ""
panel1(3).Caption = ""
panel1(4).Caption = ""

Fixgrd
grid1.rows = 1
MyAddItem

fixgrd2
grid2.rows = 1
myAddItem2

Fixgrd3
grid3.rows = 1

Handlecontrols DefineMode
On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
grid1.SetFocus
Err.Clear
If SSTab1.Tab = 0 Then SSTab1.Tab = 1
End Sub
Sub myProc(Optional sControl As String)
If ActiveControl.Name = cmdInform.Name Then
    xCode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = Me.cmdInform_rel.Name Then
    xCode.text = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 0)
    oSearchRel.Hide
    myUndo
ElseIf sControl = cmdClaim.Name Or sControl = cmdClaimFawry.Name Then
    Dim sType As String, sMsg As String
    sType = oSearchClaim.grid1.TextMatrix(oSearchClaim.grid1.Row, 0)
    sMsg = validClaim(xCode.text, myFormat(Date), sType & "", con)
    Unload oSearchClaim
    If sMsg <> "ok" Then
        MsgBox sMsg
        Exit Sub
    Else
        aRet = DocSameDay(xCode.text, sType, myFormat(Date), con)
        
        Dim oPaid As New paidfrm
        Set oPaid.myForm = Me
        oPaid.bFawry = sControl = cmdClaimFawry.Name
        
        atype = Claim_Type_Load(sType, , con)
        
        
        If (Not IsEmpty(aRet)) And retFlag(atype, "type") <> 300 Then
            MsgBox "„” ‰œ »‰›” ‰Ê⁄ «·„ÿ«·»… »‰›” «·ÌÊ„ —ﬁ„ " & aRet
            oPaid.sDoc_no = aRet
            oPaid.Show
        Else
            oPaid.bNew = True
            oPaid.sCode = xCode.text
            oPaid.sType = sType
            oPaid.Show
        End If
        Myloadgrd3
    End If
ElseIf ActiveControl.Name = cmdService.Name Then
    Dim oPaidService As New paidfrm
    Set oPaidService.myForm = Me
    oPaidService.sDoc_no = oSearchService.grid1.TextMatrix(oSearchService.grid1.Row, 0)
    Unload oSearchService
    oPaidService.Show
End If
End Sub
Private Sub myload()
xCode.text = CardTable!code & ""
xTitle.text = CardTable!Title & ""
xDesca.text = CardTable!Desca & ""
xDied.Value = IIf(CardTable!died, 1, 0)
xDrop.Value = IIf(CardTable!Drop, 1, 0)
xChair.Value = IIf(CardTable!Chair, 1, 0)
xCard.Caption = CardTable!card & ""
xCode_main.text = CardTable!code_main & ""
bCheck = False
xInstall.Value = IIf(CardTable!INSTALL, 1, 0)
bCheck = True

'xFace.Text = CardTable!Face & ""
'xFace.Text = CardTable!Face & ""
xDate_birth.text = myFormat_p(CardTable!DATE_BIRTH)
xDate_Begin.text = myFormat_p(CardTable!date_begin)

'xDate_Trans.Text = myFormat_p(CardTable!DATE_TRANS)
xSes_no.text = CardTable!SES_NO & ""
If chkFawry.Value = 1 Then
    xFawryValue.Caption = CardTable!Fawryvalue & ""
Else
    xFawryValue.Caption = ""
End If

'x.Text = CardTable!SES_NO & ""

xGender.BoundText = CardTable!GENDER & ""
xSocial.BoundText = CardTable!SOCIAL & ""
xCompany.BoundText = CardTable!Company & ""
xReligion.BoundText = CardTable!RELIGION & ""
xId_no.text = CardTable!ID_NO & ""
xAddress.text = CardTable!Address & ""
xJob_desca.text = CardTable!JOB_desca & ""
xPhone.text = CardTable!Phone & ""
xMobil.text = CardTable!Mobil & ""
'xMail.Text = CardTable!MAIL & ""
xJob.BoundText = CardTable!job & ""
xPhone_work.text = CardTable!Phone_work & ""
xJob_desca.text = CardTable!JOB_desca & ""
xDegree.BoundText = CardTable!Degree & ""
xCompany.BoundText = CardTable!Company & ""
xRegion.BoundText = CardTable!Region & ""
xapg.Value = IIf(CardTable!apg, 1, 0)

xDate_Begin.text = myFormat_p(CardTable!date_begin)
xType.BoundText = CardTable!Type & ""

aPaid = Member_Paid(xCode.text, , con)
xDate_Paid.Caption = myFormat_p(retFlag(aPaid, "Date"))
xDate1.Caption = myFormat_p(retFlag(aPaid, "Date1"))
xdate_end.Caption = myFormat_p(retFlag(aPaid, "Date2"))
xPaid_desca.Caption = retFlag(aPaid, "paid_desca") & ""
xDoc_No.Caption = retFlag(aPaid, "doc_no") & ""
xValue.Caption = retFlag(aPaid, "total") & ""
xNotes.text = CardTable!notes & ""
'xCode_main.Text = CardTable!code_main & ""
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

panel1(1).Caption = CardTable!UserName & ""
panel1(2).Caption = myFormat_p(CardTable!Time, True)
panel1(3).Caption = CardTable!UserName2 & ""
panel1(4).Caption = myFormat_p(CardTable!Time2, True)

LoadPhoto xCode.text
xDate_print.Caption = myFormat_p(CardTable!DATE_PRINT)
myLoadGrd
myloadgrd2
Myloadgrd3

On Error Resume Next
CellPos 13, 0, grid1.Cols - 1
cellPos2 13, 0, grid2.Cols - 1

loadPhoto_Append xCode.text, grid1.TextMatrix(grid1.Row, 0)

'If SSTab1.Tab = 0 Then grid1.SetFocus Else grid2.SetFocus
Err.Clear
End Sub
Private Function myreplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "Title", addstring(xTitle.text))
aInsert = AddFlag(aInsert, "Desca", addstring(xDesca.text))
aInsert = AddFlag(aInsert, "Date_birth", addstring(xDate_birth.text))
aInsert = AddFlag(aInsert, "Date_Begin", addDate(xDate_Begin.text))

aInsert = AddFlag(aInsert, "SES_NO", addstring(xSes_no.text))
aInsert = AddFlag(aInsert, "CODE_MAIN", addstring(xCode_main.text))

aInsert = AddFlag(aInsert, "Died", xDied.Value)
aInsert = AddFlag(aInsert, "[Drop]", xDrop.Value)
aInsert = AddFlag(aInsert, "[apg]", xapg.Value)
aInsert = AddFlag(aInsert, "[chair]", xChair.Value)
aInsert = AddFlag(aInsert, "Gender", addvalue(xGender.BoundText))
aInsert = AddFlag(aInsert, "Social", addvalue(xSocial.BoundText))
aInsert = AddFlag(aInsert, "Religion", addvalue(xReligion.BoundText))
aInsert = AddFlag(aInsert, "Id_no", addstring(xId_no.text))
aInsert = AddFlag(aInsert, "Address", addstring(xAddress.text))
aInsert = AddFlag(aInsert, "Phone", addstring(xPhone.text))
aInsert = AddFlag(aInsert, "Mobil", addstring(xMobil.text))
aInsert = AddFlag(aInsert, "Install", xInstall.Value)

aInsert = AddFlag(aInsert, "Job", addvalue(xJob.BoundText))
aInsert = AddFlag(aInsert, "Phone_work", addstring(xPhone_work.text))
aInsert = AddFlag(aInsert, "NOTES", addstring(xNotes.text))
aInsert = AddFlag(aInsert, "Degree", addvalue(xDegree.BoundText))
aInsert = AddFlag(aInsert, "Region", addvalue(xRegion.BoundText))
aInsert = AddFlag(aInsert, "company", addvalue(xCompany.BoundText))
aInsert = AddFlag(aInsert, "Job_desca", addstring(xJob_desca.text))
aInsert = AddFlag(aInsert, "Type", addvalue(xType.BoundText))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName & " [" & GetComputerName & "]"))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")

con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    aInsert = AddFlag(aInsert, "Code", addvalue(xCode.text))
    con.Execute addInsert(aInsert, "FILE1_10")
Else
    con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & addvalue(xCode.text))
End If
If (Row = -1 And Row2 = -1) Or Row <> -1 Then myreplaceGrd Row
If (Row = -1 And Row2 = -1) Or Row2 <> -1 Then myreplaceGrd2 Row2
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CardTimer_Timer()
Dim sCard As String
Me.MousePointer = 11
sCard = Clipboard.GetText
If sCard <> "" Then
    CardTimer.Enabled = False
    Me.MousePointer = 0
    If TimerMode = 0 Then
        nRecords = SetMemberCard(xCode.text, sCard)
    ElseIf grid1.TextMatrix(grid1.Row, 0) <> "" Then
        nRecords = SetRelCard(xCode.text, grid1.TextMatrix(grid1.Row, 0), sCard)
    End If
    If nRecords > 0 Then Inform " „ «÷«›… —ﬁ„ «·ﬂ«—‰ÌÂ »‰Ã«Õ"
    If TimerMode = 0 Then
        xCard.Caption = sCard
    Else
        myLoadGrd
    End If
End If
End Sub

Private Sub chkFawry_Click()
myUndo
End Sub

Private Sub chkFawryHidden_Click()
myUndo
End Sub

Private Sub chkHideInstall_Click()
myUndo
End Sub
Private Sub chkInstall_Click()
'xInstall.Enabled = chkInstall.Value = 1
myUndo
End Sub

Private Sub cmdAdd_Click()
mydefine
xCode.SetFocus
End Sub

Private Sub cmdPrintAll_Click()
End Sub

Private Sub cmdDeliverAll_Click()
If MsgBox(" ”·Ì„ ﬂ· «·ﬂ«—‰ÌÂ«  «·€Ì— „”·„…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub

con.BeginTrans
On Error GoTo myerror
con.Execute "update file4_10 set " & _
            " file4_10.date_delivery = " & DateSq(Date) & _
            " where [year] = " & sSeason & _
            " and member = " & xCode.text
con.CommitTrans
Inform " „  ”·Ì„ «·ﬂ«—‰ÌÂ«  »‰Ã«Õ"
myloadgrd2
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub cmdDocument_Click()
Dim oPhotos As New documentsfrm
oPhotos.sCode = xCode.text
oPhotos.Show 1
End Sub

Private Sub cmdService_Click()
cWhere = "vw_last_paid_service.code = " & xCode.text
If chkCurrent.Value = 1 Then cWhere = cWhere & " and vw_last_paid_service.year_code = " & sSeason
ServiceLookup Me, oSearchService, , False, False, cmdService.TagVariant
End Sub

Private Sub cmdCard_Click()
cardsfrm.sCode = xCode.text
cardsfrm.ntype = 0
cardsfrm.Show 1
myUndo
'cmdCardTrans_Click
End Sub
Private Sub cmdCardReader_Click()
Dim sPath As String, sCard As String, nEdited As Integer
Clipboard.Clear
sPath = sPath_App & "\CardReader\CardReader.exe"
RunIt sPath, vbNormalFocus
TimerMode = 0
CardTimer.Enabled = True
End Sub

Private Sub cmdCardTrans_Click()
Dim con2 As New ADODB.Connection, cInsert As String
sMsg = openCon(con2, CreateConStr2)
If sMsg <> "ok" Then
    MsgBox sMsg
    Exit Sub
End If
cInsert = SendCard(xCode.text, , con, con2)
If cInsert <> "" Then
    con2.BeginTrans
    On Error GoTo myerror
    con2.Execute cInsert
    con2.CommitTrans
    MsgBox " „ ‰ﬁ· «·»Ì«‰«  ··»Ê«»… »‰Ã«Õ"
End If
closeCon con2
Exit Sub
myerror:
MsgBox Err.Description
con2.RollbackTrans
Err.Clear
End Sub
Private Function ValidFawry() As Boolean
Dim nValue As Double
nValue = GetField("select [dbo].[fawry_acount](" & addvalue(xCode.text) & ")", con)
If nValue <= 0 Then
    MsgBox ("·Ì” ··⁄„Ì· Õ”«» ›Ê—Ì ·⁄„· „ÿ«·»… ›Ê—Ì")
    Exit Function
End If
ValidFawry = True
End Function

Private Sub cmdClaim_Click()
    claim_LookupAll Me, oSearchClaim, cmdClaim.Name
End Sub

Private Sub cmdClaimFawry_Click()
If ValidFawry Then
    claim_LookupAll Me, oSearchClaim, cmdClaimFawry.Name
End If

End Sub

Private Sub cmdCompany_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xCompany.BoundText
oFlagfrm.sTable = "company_CODES"
oFlagfrm.sCaption = "«·‘—ﬂ…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA4.Recordset = myRecordSet("select * from company_Codes", con)
xCompany.BoundText = sBound
If Not xCompany.MatchedWithList Then xCompany.BoundText = ""
End Sub

Private Sub cmdDegree_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xDegree.BoundText
oFlagfrm.sTable = "Degree_CODES"
oFlagfrm.sCaption = "«·ÊŸÌ›…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
data6.Recordset.Requery
xDegree.BoundText = sBound
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub

Private Sub CmdDel_Click()
On Error GoTo myerror
If xDrop.Value = 0 Then
    MsgBox "«·⁄÷Ê ·Ì” ”«ﬁÿ ⁄÷ÊÌ…"
    Exit Sub
End If
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If grid1.rows > 2 Then
        MsgBox "«·⁄÷Ê ·Â  Ê«»⁄ ÌÃ» Õ–›Â„ «Ê·«"
        Exit Sub
    End If
    con.BeginTrans
    con.Execute "Delete  From FILE1_10 Where code = " & xCode.text & " AND FILE1_10.[DROP] = 1", nDelete
    con.CommitTrans
    If nDelete = 0 Then
        MsgBox "·„ Ì „ «·‰Ÿ«„ „‰ Õ–› «·⁄÷Ê"
        Exit Sub
    End If
    DeletePhoto xCode.text
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

Private Sub cmdDelCard_Click()
End Sub

Private Sub cmdDelCardRel_Click()

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
If Trim(xCode.text) = "" Then Exit Sub
If xCode.Tag = DefineMode Then Exit Sub
Set fs = CreateObject("Scripting.FileSystemObject")
Dim cFile As String, cNewFile As String
On Error GoTo myerror
Common1.FileName = ""
Common1.InitDir = App.Path & "\PICT"
Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
Common1.ShowOpen
If Common1.FileTitle <> "" Then
    cFile = Common1.FileName
    If cFile <> "" Then
        fs.CopyFile cFile, retPhoto(xCode.text)
    End If
    myload
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CMDFIXNO_Click()
Dim loctable As New ADODB.Recordset
Dim bFound As Boolean
loctable.Open "Select * from file1_10 where not ses_no is null order by code", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = nRecord
    sSes_no = ""
    If (Not IsDate(loctable!SES_NO)) Then
        bFound = False
        For I = Len(loctable!SES_NO & "") To 1 Step -1
            If IsNumeric(Mid(loctable!SES_NO, I, 1)) Then
                bFound = True
                sSes_no = Mid(loctable!SES_NO, I, 1) + sSes_no
            ElseIf bFound Then
                Exit For
            End If
        Next
    End If
    If Len(sSes_no) >= 4 Then
        con.Execute "update file1_10 set file1_10.code_main = " & addvalue(sSes_no) & " where code = " & loctable!code
    End If
    loctable.MoveNext
Loop
End Sub

Private Sub CmdNext_Click()
openCardTable xCode.text, ">"
If CardTable.EOF Then openCardTable xCode.text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.text, "<"
If CardTable.EOF Then openCardTable xCode.text, "="
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
Private Sub CmdInform_Click()
MemberLookupAll Me, oSearch, cFilter
End Sub
Private Sub cmdInform_rel_Click()
relLookupAll Me, oSearchRel
End Sub

Private Sub cmdJob_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xJob.BoundText
oFlagfrm.sTable = "JOB_CODES"
oFlagfrm.sCaption = "«·ÊŸÌ›…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA5.Recordset = myRecordSet("select * from JOB_Codes", con)
xJob.BoundText = sBound
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub

Private Sub cmdSection_Click()
Dim oFlagfrm As New flag_mainfrm, sBoundText As String
sBoundText = xRegion.BoundText
oFlagfrm.sTable = "SECTION_CODES"
oFlagfrm.sCaption = "«·«œ«—… «·⁄«„…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA7.Recordset = myRecordSet("select * from section_Codes", con)
xRegion.BoundText = sBoundText
If Not xRegion.MatchedWithList Then xRegion.BoundText = ""
End Sub
Private Sub cmdQual_Click()
Dim myPublic(5)
nCode = xQUAL_CODE.BoundText
myPublic(0) = "Qual_codes"
myPublic(1) = "Code"
myPublic(2) = "Desca"
myPublic(3) = "ﬂÊœ «·„ƒÂ·"
myPublic(4) = "«·„ƒÂ·"
myPublic(5) = "«ﬂÊ«œ «·„ƒÂ·« "
FlagFrm.bedit = True
FlagFrm.myPublic = myPublic
FlagFrm.Show 1
DATA3.Refresh
xQUAL_CODE.BoundText = nCode
If Not xQUAL_CODE.MatchedWithList Then xQUAL_CODE.BoundText = ""
End Sub

Private Sub cmdRegion_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xRegion.BoundText
oFlagfrm.sTable = "REGION_CODES"
oFlagfrm.sCaption = "«· ﬁ”Ì„ «·«œ«—Ì"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
DATA7.Recordset.Requery
xRegion.BoundText = sBound
If Not xRegion.MatchedWithList Then xRegion.BoundText = ""
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
'openCardTable
myUndo
End Sub
Private Sub cmdScan_Click()
Scanfrm.sCode = xCode.text
Scanfrm.Show 1
If validPhoto(retPhoto(xCode.text)) Then xMemberPhoto.Picture = LoadPicture(retPhoto(xCode.text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.text, grid1.TextMatrix(grid1.Row, 0)))
End If
myload
End Sub

Private Sub cmdStatus_Click()
'Dim oFlagfrm As New flag_mainfrm, sBound As String
'sBound = xStatus.BoundText
'oFlagfrm.sTable = "Status_CODES"
'oFlagfrm.sCaption = "«·Õ«·…"
'oFlagfrm.nZero = -1
'oFlagfrm.bEdit = True
'oFlagfrm.Show 1
'Set DATA4.Recordset = myRecordSet("select * from Status_Codes", con)
'xStatus.BoundText = sBound
'If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub

Private Sub cmdSendAll_Click()
DoorCardSend.ntype = 0
DoorCardSend.Show 1
End Sub

Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub
Private Sub cmdScan2_Click()
nPhoto = 0
ScanImage
On Error Resume Next
If validPhoto(retPhoto(xCode.text)) Then xMemberPhoto.Picture = LoadPicture(retPhoto(xCode.text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.text, grid1.TextMatrix(grid1.Row, 0)))
End If
Err.Clear
End Sub
Private Sub Command1_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, I As Long, cFile As String, nRecordcount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
'loctable.Open "select * from file1_10 where NEWDATA = true", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.Open "select * from file1_10", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordcount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    I = I + 1
    Me.Caption = cCaption & I & " from " & nRecordcount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, retPhoto(loctable!code)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub

Private Sub Command10_Click()
'Dim con2 As New ADODB.Connection, loctable As New ADODB.Recordset
'cFile = App.Path & "\temp.txt"
'openCon con2, LoadConString(, "Olympic2")
'loctable.Open "SELECT [NO] FROM MEMBERS WHERE DEAD = 1", con2, adOpenStatic, adLockReadOnly, adCmdText
'Open cFile For Output As #1   ' Open file for output.
'Do Until loctable.EOF
'    aInsert = AddFlag(Empty, "DIED", "1")
'    cInsert = addUpdate(aInsert, "FILE1_10", "[CODE] = " & loctable!NO) & ";"
'    Print #1, cInsert
'    loctable.MoveNext
'Loop
'MsgBox "done"
End Sub

Private Sub addPaid()
Dim conMdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror
conMdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM [claim_values]", conMdb, adOpenStatic, adLockReadOnly, adCmdText

Dim nRecordcount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordcount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordcount
    Dim aSep As Variant
    aSep = Split(loctable!CODE_ALL, "-")
    aInsert = AddFlag(Empty, "CODE", aSep(1))
    aInsert = AddFlag(aInsert, "MEMBER_SPLIT", aSep(0))
    aInsert = AddFlag(aInsert, "CODE_ALL", loctable!CODE_ALL)
    aInsert = AddFlag(aInsert, "MEMBER", addvalue(loctable!member))
    aInsert = AddFlag(aInsert, "MEMBERID", addstring(loctable!MEMBERID))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca & ""))
    aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(Format(loctable!DATE_BIRTH, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "RELATION", addvalue(loctable!RELATION & ""))
    aInsert = AddFlag(aInsert, "SECTION", addvalue(loctable!Section & ""))
    aInsert = AddFlag(aInsert, "GENDER", addvalue(loctable!GENDER & ""))
    aInsert = AddFlag(aInsert, "union_reg", addstring(loctable!Union_reg & ""))
    aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!notes & ""))
    aInsert = AddFlag(aInsert, "JOB_CODE", addstring(loctable!JOB_CODE & ""))
    aInsert = AddFlag(aInsert, "PHOTO_CODE", addstring(loctable!PHOTO_CODE & ""))
    con.Execute addInsert(aInsert, "FILE1_11")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conMdb.Close
Set conMdb = Nothing
MsgBox "Done"
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Command11_Click()
'Dim loctable As New ADODB.Recordset
'loctable.Open "SELECT * FROM FILE6_20H WHERE DATE >= '2018-01-01' AND DATE <= '2019-06-30'", con, adOpenStatic, adLockReadOnly, adCmdText
'nRecordcount = loctable.RecordCount
'Do Until loctable.EOF
'    nRecord = nRecord + 1
'    Me.Caption = "Record " & nRecord & " from " & nRecordcount
'    cString = fixClaim(loctable!DOC_NO)
'    con.Execute cString
'    loctable.MoveNext
'Loop
'MsgBox "done..."
End Sub

Private Sub Command4_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, I As Long, cFile As String, nRecordcount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
loctable.Open "select * from file1_11", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordcount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    I = I + 1
    Me.Caption = cCaption & I & " from " & nRecordcount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, RetAppendPhoto(loctable!member, loctable!code)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub
Private Sub Command5_Click()
'Dim loctable As ADODB.Recordset, cString As String, cPhoto As String
'Me.MousePointer = 11
'Dim fs As New FileSystemObject
'cString = "select * from FILE1_10 ORDER BY CODE "
'Set loctable = New ADODB.Recordset
'loctable.Open cString, con, adOpenStatic, adLockReadOnly
'Do Until loctable.EOF
'    I = I + 1
'    Me.Caption = I & " From "
'    cPhoto = App.Path & "\photo_s\" & loctable!CODE & ".jpg"
'    If validPhoto(cPhoto) Then
'        fs.MoveFile cPhoto, retPhoto(loctable!CODE)
'    End If
'    loctable.MoveNext
'Loop
'
'cString = "select * from FILE1_11 ORDER BY MEMBER,CODE"
'Set loctable = New ADODB.Recordset
'loctable.Open cString, con, adOpenStatic, adLockReadOnly
'Do Until loctable.EOF
'    cPhoto = App.Path & "\photo_s\" & loctable!MEMBER & "-" & loctable!CODE & ".jpg"
'    If validPhoto(cPhoto) Then
'        fs.MoveFile cPhoto, RetAppendPhoto(loctable!MEMBER, loctable!CODE)
'        I = I + 1
'        Me.Caption = I & " From "
'    End If
'    loctable.MoveNext
'Loop
'MsgBox " „ ”Õ» «·’Ê— »‰Ã«Õ"
End Sub

Private Sub Command6_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
Set f = fs.GetFolder(App.Path & "\photo_fix")
Set fc = f.Files
nCount = fc.Count
Dim cString As String, I As Long
bCrypt = True
For Each f1 In fc
    I = I + 1
    Me.Caption = I
    If InStr(1, LCase(App.Path & "\photo_fix\" & f1.Name), "jpg") <> 0 Then
        cFile = retPhoto(Replace(LCase(f1.Name), ".jpg", ""))
        If cFile <> "" Then
            fs.CopyFile App.Path & "\photo_fix\" & f1.Name, cFile
        Else
           con.Execute "INSERT INTO TEST(CODE) " & _
                        "VALUES(" & _
                        addstring(f1.Name) & _
                        ")"
        End If
    End If
Next
MsgBox "done..."
End Sub

Private Sub Command511_Click()
'Dim loctable As New ADODB.Recordset
'con.Execute "UPDATE FAWRY_TRANS SET FAWRY_TRANS.BILL_AC_NO = FAWRY_TRANS.BILL_AC_NO"
con.Execute "UPDATE FILE1_10 SET FILE1_10.TAXABLE = 1 WHERE FILE1_10.DATE_BEGIN > '2016-08-09'"
End Sub

Private Sub Command7_Click()
Dim loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror

loctable.Open "SELECT * FROM FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdText

Dim nRecordcount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordcount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant, bNoCard As Boolean
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordcount
    bNoCard = Not validPhoto(retPhoto(loctable!code))
    aInsert = AddFlag(Empty, "nocard", IIf(bNoCard, "1", "0"))
    con.Execute addUpdate(aInsert, "FILE1_10", "CODE = " & loctable!code)
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
Exit Sub
myerror:
MsgBox Err.Description
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xCode.Tag = LoadMode Then
        If SSTab1.Tab = 2 Then
            grid1.SetFocus
        Else
            grid2.SetFocus
        End If
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

AddOtherType

On Error Resume Next
CSubclass1.SubClassMe SSTab1.hwnd, 0, , vbWhite       '//--- Begin SubClassing
Err.Clear

cRelStr = StrList2("Select Code,Desca From relation_codes order by desca")
cGenderStr = StrList2("Select Code,Desca From gender_codes order by Code")

Set DATA7.Recordset = myRecordSet("select * from Region_Codes", con)
Set xRegion.RowSource = DATA7
xRegion.ListField = "Desca"
xRegion.BoundColumn = "Code"

Set DATA1.Recordset = myRecordSet("select * from Gender_Codes", con)
Set xGender.RowSource = DATA1
xGender.ListField = "Desca"
xGender.BoundColumn = "Code"

Set DATA2.Recordset = myRecordSet("select * from religion_Codes", con)
Set xReligion.RowSource = DATA2
xReligion.ListField = "Desca"
xReligion.BoundColumn = "Code"

Set DATA3.Recordset = myRecordSet("select * from social_Codes", con)
Set xSocial.RowSource = DATA3
xSocial.ListField = "Desca"
xSocial.BoundColumn = "Code"

Set DATA4.Recordset = myRecordSet("select * from company_Codes", con)
Set xCompany.RowSource = DATA4
xCompany.ListField = "Desca"
xCompany.BoundColumn = "Code"

Set DATA5.Recordset = myRecordSet("select * from Job_Codes", con)
Set xJob.RowSource = DATA5
xJob.ListField = "Desca"
xJob.BoundColumn = "Code"

Set data6.Recordset = myRecordSet("select * from Degree_Codes", con)
Set xDegree.RowSource = data6
xDegree.ListField = "Desca"
xDegree.BoundColumn = "Code"

'Set data8.Recordset = myRecordSet("select * from reason_Codes", con)
'Set xReason.RowSource = data8
'xReason.ListField = "Desca"
'xReason.BoundColumn = "Code"

Set data9.Recordset = myRecordSet("select * from type_Codes", con)
Set xType.RowSource = data9
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set grid1.DataSource = DATA11
Set grid2.DataSource = DATA12
Set grid3.DataSource = DATA13

bedit = Not retFlag(aSec, "INFORM")
HandleFirst
Fixgrd
'openCardTable
If nTab <> 0 Then SSTab1.Tab = nTab
myUndo
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then Exit Sub
Dim sPath As String, sCard As String, nEdited As Integer
Clipboard.Clear
sPath = sPath_App & "\CardReader\CardReader.exe"
RunIt sPath, vbNormalFocus
TimerMode = 1
CardTimer.Enabled = True
End Sub

Private Sub grid3_DblClick()
If xCode.Tag = DefineMode Then Exit Sub
If Not ValidNum(xCode.text) Then Exit Sub
If grid3.TextMatrix(grid3.Row, 0) <> "" And grid3.Row > 0 Then
    paidfrm.sDoc_no = grid3.TextMatrix(grid3.Row, 0)
    paidfrm.Show
End If
End Sub

Private Sub SSCommand1_Click()
End Sub

Private Sub xCard_DblClick()
If xCard.Caption = "" Then Exit Sub
If MsgBox("Õ–› —ﬁ„ «·ﬂ«—‰ÌÂ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
xCard.Caption = ""
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
If Not ValidNum(xCode.text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If Trim(xDesca.text) = "" Then
    MsgBox "√”„ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not xType.MatchedWithList Then
    MsgBox "›∆… «·⁄÷ÊÌ… €Ì— „”Ã·…", , systemName
    Exit Function
End If

If Not IsDate(xDate_birth.text) Then
    MsgBox " «—ÌŒ «·„Ì·«œ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not IsDate(xDate_Begin.text) Then
    MsgBox " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not xGender.MatchedWithList Then
    MsgBox "«·‰Ê⁄ €Ì— „”Ã·", , systemName
    Exit Function
End If

If Age(myFormat(xDate_birth.text)) < 21 Then
    MsgBox "«·”‰ " & Age(myFormat(xDate_birth.text)) & " ”‰Ê« "
    Exit Function
End If

sField = myField("select code from file1_10 where desca = " & MyParn(xDesca.text) & " and code <> " & MyParn(xCode.text), "code", con) & ""
If sField <> "" Then
    MsgBox "«·«”„ „”Ã· „‰ ﬁ»· ··⁄÷ÊÌ… —ﬁ„ " & sField
    Exit Function
End If


For I = 1 To grid1.rows - 2
    If grid1.TextMatrix(I, 1) = "1" Then
        If IsDate(myFormat(grid1.TextMatrix(I, 5))) Then
            If Age(myFormat(grid1.TextMatrix(I, 5))) < 21 Then
                MsgBox "”‰ «·“ÊÃ… " & Age(myFormat(grid1.TextMatrix(I, 5))) & " ”‰Ê« "
                Exit Function
            End If
        End If
    End If
Next
'If bIgMsg Then
'    For I = 1 To grid1.rows - 2
'        If Not ValidInt(grid1.TextMatrix(I, 0)) Then
'            MsgBox "ﬂÊœ «· «»⁄ €Ì— „”Ã·"
'            Exit Function
'        End If
'
'
'        If Not ValidInt(grid1.TextMatrix(I, 1)) Then
'            MsgBox "‰Ê⁄ «· »⁄Ì… €Ì— „”Ã·…"
'            Exit Function
'        End If
'    Next
'End If
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveText Me, , Array(xcode1.Name, xCode2.Name)
CardTable.Close
CaseTable.Close
Set CardTable = Nothing
Set memberfrm = Nothing
End Sub
Private Sub PrintMembers()
Dim cString As String, temptable As New ADODB.Recordset, loctable As New ADODB.Recordset

contemp.Execute "delete  from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

cString = "SELECT FILE1_10.*, FILE1_11.MEMBER, FILE1_11.DESCA AS DESCA_REL, FILE1_11.DATE_BIRTH AS DATE_BIRTH_REL, FILE1_11.PRINT_DATE AS PRINT_DATE_REL, REL_CODES.DESCA AS REL_CODE_DESCA" & _
          " FROM (FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER) LEFT JOIN REL_CODES ON FILE1_11.RELATION = REL_CODES.CODE"

If IsNumeric(xcode1.text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.text), " >= ", " = ") & xcode1.text
End If

If IsNumeric(xCode2.text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.text
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Do Until loctable.EOF
    temptable.AddNew
    temptable!val1 = loctable!code
    temptable!str1 = ArbString(loctable!code)
    temptable!str2 = loctable!vip
    temptable!str3 = loctable!Name
    temptable!str4 = loctable!Title
    temptable!str5 = loctable!Address
    If Not IsNull(loctable!Degree) Then
        temptable!str6 = GetField("select desca from degree_Codes where code = " & UnCodeSerial(CardTable!Degree, 71))
    End If
    temptable!str7 = loctable!Address
    temptable!str8 = loctable!phone1
    temptable!str9 = loctable!Mobil
    temptable!str10 = loctable!Union
    
    temptable!str11 = TurnValue(ArbString(Format(loctable!DATE_BIRTH, "yyyy/mm/dd")))
    temptable!str12 = TurnValue(ArbString(loctable!receipt & ""))
    temptable!str13 = TurnValue(ArbString(Format(loctable!Print_date, "yyyy/mm/dd")))
    
    temptable!val2 = loctable!member
    temptable!str16 = loctable!Desca_rel
    temptable!str17 = loctable!REL_CODE_DESCA
    temptable!str18 = TurnValue(ArbString(Format(loctable!Print_date_rel, "yyyy/mm/dd")))
    temptable!str19 = TurnValue(ArbString(Format(loctable!DATE_BIRTH_rel, "yyyy/mm/dd")))
    'temptable!Val3 = retPaid(locTable!CODE)
    temptable.Update
    loctable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    temptable.Requery
    con.BeginTrans
    con.CommitTrans
    REPORT1.ReportFileName = MainPath & "\rpt\Member_data.rpt"
    REPORT1.DataFiles(0) = cTempPath
    REPORT1.Action = 1
End If
Set temptable = Nothing
Set loctable = Nothing
End Sub
Private Sub CalcTotals()
Dim nValue
For I = 1 To grid1.rows - 2
    nValue = Val(grid1.TextMatrix(I, 7)) + nValue
Next
If xDied.Value = 0 Then nValue = nValue + nMemValue
xTotal.Caption = Format(nValue, "fixed")
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
If chkFawry.Value = 1 Then
    cString = "SELECT TOP 1 FILE1_10.*,FAWRY_BALANCE.BALANCE AS FAWRYVALUE  FROM FILE1_10 INNER JOIN FAWRY_BALANCE ON FILE1_10.CODE = FAWRY_BALANCE.CODE"
Else
    cString = "SELECT TOP 1 FILE1_10.*  FROM FILE1_10"
End If
If pCode <> "" Then cWhere = "FILE1_10.CODE " & pSign & addvalue(pCode)

cFilter = ""
If chkFawryHidden.Value = 1 Then cFilter = cFilter & turn(cFilter, " and ") & "FILE1_10.CODE IN(SELECT MEMBER FROM FILE1_15)"
If sCode <> "" Then cFilter = "FILE1_10.CODE = " & addvalue(sCode)
If chkInstall.Value = 1 Then cFilter = cFilter & turn(cFilter, " AND ") & "FILE1_10.TAXABLE = 1"
If chkHideInstall.Value = 1 Then cFilter = cFilter & turn(cFilter, " AND ") & "FILE1_10.INSTALL = 0"

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
If chkFawry.Value = 1 Then
    cString = cString & " FROM FILE1_10 INNER JOIN FAWRY_BALANCE ON FILE1_10.CODE = FAWRY_BALANCE.CODE " & turn(cFilter, " WHERE ") & cFilter
Else
    cString = cString & " FROM FILE1_10 " & turn(cFilter, " WHERE ") & cFilter
End If

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If ValidNum(pCode) Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
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
If Not IsNumeric(xCode.text) Then Exit Sub
If nPhoto = 0 And xCode.text Then
    ReplaceFromImage Image, retPhoto(xCode.text)
Else
    If nPhoto <= grid1.rows - 1 Then
        If IsNumeric(grid1.TextMatrix(nPhoto, 0)) Then
            ReplaceFromImage Image, retPhoto(xCode.text & "-" & grid1.TextMatrix(nPhoto, 0))
        End If
    End If
End If
nPhoto = nPhoto + 1
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
Private Sub grid1_KeyPress(KeyAscii As Integer)
With grid1
If KeyAscii = 13 And (.Col <> 1 And .Col <> 2) Then KeyAscii = 0
End With
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
With grid1
    If KeyCode = 46 And .Row <> .rows - 1 And bEditRecord Then
        If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbDefaultButton2 + vbOKCancel) Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                Dim fs As New FileSystemObject
                If Trim(.TextMatrix(Row, 0)) <> "" Then
                    DeletePhoto xCode.text, .TextMatrix(.Row, 0)
                End If
                con.BeginTrans
                con.Execute "Delete  from file1_11 where id = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            myRemove .Row
            grid1_EnterCell
            On Error Resume Next
            grid1.SetFocus
            Err.Clear
            loadPhoto_Append xCode.text, grid1.TextMatrix(grid1.Row, 0)
        End If
    ElseIf KeyCode = 13 Then
        CellPos KeyCode, .Row, .Col
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And (Col <> 1 And Col <> 2) Then CellPos KeyCode, Row, Col
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo myerror
With grid1
If Not MYVALID Then
    On Error Resume Next
    .SetFocus
    Err.Clear
    myLoadGrd
    If Row < .rows - 1 Then
        .Select Row, Col
    Else
        CellPos 13, .rows - 2, .Cols - 1
    End If
    Exit Sub
End If
If Not validRow(Row) Then Exit Sub
If Row = .rows - 1 Then
    MyAddItem
End If
'Calctotals
If myreplace(Row) Then
    If xCode.Tag = DefineMode Then
        Handlecontrols LoadMode
        myLoadGrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myLoadGrd
    End If
Else
    myLoadGrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myLoadGrd
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not ValidNum(.TextMatrix(Row, 0)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 4)) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 5)) Then Exit Function
If Not IsDate(.TextMatrix(Row, 6)) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow < 1 Then Exit Sub
On Error GoTo myerror
If OldRow <> NewRow Then
    loadPhoto_Append xCode.text, .TextMatrix(NewRow, 0)
End If
End With
Exit Sub
myerror:
xAppendPhoto.Picture = LoadPicture("")
End Sub
Private Sub grid1_EnterCell()
With grid1
If (.Col = 0 And Trim(.TextMatrix(.Row, .Cols - 1)) <> "") Or .Col = 10 Or .Col = 11 Or .Col = 12 Then
    .Editable = flexEDNone
Else
    .Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
grid1_EnterCell
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If OldRow < 1 Then Exit Sub
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Not ValidNum(.EditText) Then
        If .Row = .rows - 1 Then Exit Sub
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        nFound = FoundOtheritem(grid1, Row, 0, Trim(.EditText))
        If nFound <> -1 Then
            MsgBox "«·ﬂÊœ „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
            Exit Sub
        End If
    End If
ElseIf Col = 2 Then
    If Trim(.EditText) = "" Then
        MsgBox "«·ﬁ—«»… €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 4 Then
    If Trim(.EditText) = "" Then
        MsgBox "«·«”„ €Ì— „”Ã·"
        Cancel = True
    End If
ElseIf Col = 5 Or Col = 6 Then
    If (Not IsDate(.EditText)) Then
        Cancel = True
    Else
        .EditText = myFormat_p(.EditText)
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·—ﬁ„|" & "«·ﬁ—«»…|" & "«·‰Ê⁄|" & "«·’›…|" & "«·«”„|" & " «—ÌŒ «·„Ì·«œ|" & " «—ÌŒ «·⁄÷ÊÌ…|" & "„·«ÕŸ« |" & "„ ÕœÌ «⁄«ﬁ…|" & "„ ŒÿÌ «·”‰|" & "«·”‰|" & " «—ÌŒ «·ÿ»«⁄…|" & "—ﬁ„ «·ﬂ«—‰ÌÂ|"
.ColWidth(0) = 800
.ColWidth(1) = 1800
.ColWidth(2) = 1200
.ColWidth(3) = 1500
.ColWidth(4) = 3500
.ColWidth(5) = 1250
.ColWidth(6) = 1250
.ColWidth(7) = 2300
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 800
.ColWidth(11) = 1300
.ColWidth(12) = 1400
.ColComboList(12) = "..."
.ColDataType(8) = flexDTBoolean
.ColDataType(9) = flexDTBoolean
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColComboList(1) = cRelStr
.ColComboList(2) = cGenderStr
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid1
KeyCode = 0
If Col < .Cols - 2 Then
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
Private Sub MyAddItem()
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
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.rows - 2, Row)
        aInsert = AddFlag(Empty, "CODE", addvalue(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "RELATION", addvalue(grid1.TextMatrix(I, 1)))
        aInsert = AddFlag(aInsert, "GENDER", addvalue(grid1.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(grid1.TextMatrix(I, 5)))
        aInsert = AddFlag(aInsert, "DATE_BEGIN", addDate(grid1.TextMatrix(I, 6)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(I, 7)))
        aInsert = AddFlag(aInsert, "HANDI", IIf(mRound(grid1.TextMatrix(I, 8)) = 0, "0", "1"))
        aInsert = AddFlag(aInsert, "PENDING", IIf(mRound(grid1.TextMatrix(I, 9)) = 0, "0", "1"))
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            aInsert = AddFlag(aInsert, "MEMBER", addvalue(xCode.text))
            con.Execute addInsert(aInsert, "FILE1_11")
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myLoadGrd()
With grid1
Dim cString As String
cString = "SELECT FILE1_11.CODE,FILE1_11.RELATION,FILE1_11.GENDER,FILE1_11.TITLE,FILE1_11.DESCA,CONVERT(VARCHAR(10),FILE1_11.DATE_BIRTH,111),CONVERT(VARCHAR(10),FILE1_11.DATE_BEGIN,111),FILE1_11.NOTES,FILE1_11.HANDI,FILE1_11.PENDING,dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(sDate_Season) & "),CONVERT(VARCHAR(10),FILE1_11.DATE_PRINT,111) ,FILE1_11.CARD,FILE1_11.ID " & _
          " FROM FILE1_11"
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.text
cString = cString & " ORDER BY FILE1_11.CODE"
Set DATA11.Recordset = myRecordSet(cString, con)
MyAddItem
Fixgrd
End With
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
If grid1.rows > 2 Then
    grid1.TextMatrix(grid1.rows - 1, 0) = grid1.TextMatrix(grid1.rows - 2, 0) + 1
ElseIf grid1.rows = 2 Then
    grid1.TextMatrix(grid1.rows - 1, 0) = 1
End If
End Sub
Private Function FoundOtheritem(grid1 As Variant, nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function

Private Sub xDrop_Click()
Handlecontrols LoadMode
End Sub

Private Sub xFawry_Hidden_Click()
myUndo
End Sub
Private Sub xFawryValue_DblClick()
If IsDate(grid2.TextMatrix(grid2.rows - 1, 0)) And Val(xFawryValue.Caption) <> 0 Then
    grid2.TextMatrix(grid2.rows - 1, 1) = xFawryValue.Caption
    myAddItem2
End If
End Sub

Private Sub xInstall_Click()
If bCheck Then cmdSave_Click
End Sub

Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_Begin_GotFocus()
myGotFocus xDate_Begin
End Sub
Private Sub xDate_Begin_LostFocus()
myLostFocus xDate_Begin
myValidDate xDate_Begin
End Sub
Private Sub xDate_Join_GotFocus()
myGotFocus xDate_Join
End Sub
Private Sub xDate_Join_LostFocus()
myLostFocus xDate_Join
myValidDate xDate_Join
End Sub
Private Sub xDate_Trans_GotFocus()
myGotFocus xDate_Trans
End Sub
Private Sub xDate_Trans_LostFocus()
myLostFocus xDate_Trans
myValidDate xDate_Trans
End Sub
Private Sub xSes_No_GotFocus()
myGotFocus xSes_no
End Sub
Private Sub xSes_No_LostFocus()
myLostFocus xSes_no
End Sub
Private Sub xReason_GotFocus()
myGotFocus xReason
End Sub
Private Sub xReason_LostFocus()
myLostFocus xReason
If Not xReason.MatchedWithList Then xReason.BoundText = ""
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xFace_GotFocus()
myGotFocus xFace
End Sub
Private Sub xFace_LostFocus()
myLostFocus xFace
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
Private Sub xDegree_GotFocus()
myGotFocus xDegree
End Sub
Private Sub xDegree_LostFocus()
myLostFocus xDegree
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub
Private Sub xRegion_GotFocus()
myGotFocus xRegion
End Sub
Private Sub xRegion_LostFocus()
myLostFocus xRegion
If Not xRegion.MatchedWithList Then xRegion.BoundText = ""
End Sub
Private Sub xJob_desca_GotFocus()
myGotFocus xJob_desca
End Sub
Private Sub xJob_desca_LostFocus()
myLostFocus xJob_desca
End Sub
Private Sub xJob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xJob_LostFocus()
myLostFocus xJob
If Not xJob.MatchedWithList Then xJob.BoundText = ""
End Sub
Private Sub xCompany_GotFocus()
myGotFocus xCompany
End Sub
Private Sub xCompany_LostFocus()
myLostFocus xCompany
If Not xCompany.MatchedWithList Then xCompany.BoundText = ""
End Sub
Private Sub Grid2_KeyPress(KeyAscii As Integer)
With grid2
If KeyAscii = 13 Then KeyAscii = 0
End With
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
With grid2
    If KeyCode = 46 And bEditRecord Then
        If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.BeginTrans
                con.Execute "Delete  from file4_10 where id = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            .RemoveItem .Row
            Grid2_EnterCell
        End If
    ElseIf KeyCode = 13 Then
        cellPos2 KeyCode, .Row, .Col
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd2
End Sub
Private Sub Grid2_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then cellPos2 KeyCode, Row, Col
End Sub
Private Sub Grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid2
If Not MYVALID Then
    On Error Resume Next
    .SetFocus
    Err.Clear
    myloadgrd2
    If Row < .rows - 1 Then
        .Select Row, Col
    Else
        cellPos2 13, .rows - 2, .Cols - 1
    End If
    Exit Sub
End If
If Not validRow2(Row) Then Exit Sub
'If Row = .rows - 1 Then
'    myAddItem2
'End If

If myreplace(, Row) Then
    If xCode.Tag = DefineMode Then
        myUndo
    End If
    If .TextMatrix(Row, .Cols - 1) = "" Then
         myloadgrd2
        .ShowCell .rows - 1, 0
    End If
Else
    myloadgrd2
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd2
End Sub
Private Function validRow2(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid2
'If Not IsDate(.TextMatrix(Row, 0)) Then Exit Function
'If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow2 = True
End Function
Private Sub Grid2_EnterCell()
With grid2
If grid2.Col = 4 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid2_GotFocus()
Grid2_EnterCell
End Sub
Private Sub Grid2_Validate(Cancel As Boolean)
With grid2
If OldRow < 1 Then Exit Sub
If (Not validRow2(.Row)) And .Row <> .rows - 1 And .TextMatrix(.Row, .Cols - 1) = "" Then
    myRemove .Row
End If
End With
End Sub
Private Sub Grid2_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid2
If Col = 4 Then
    If (Not IsDate(.EditText)) And Trim(.EditText) <> "" Then
        MsgBox " «—ÌŒ €Ì— „”Ã·"
        Cancel = True
    Else
        .EditText = myFormat_p(.EditText)
    End If
End If
End With
End Sub
Private Sub fixgrd2()
With grid2
    .FormatString = "«·ﬂÊœ|" & "«·«”„|" & "‰Ê⁄ «·⁄÷ÊÌ…|" & " «—ÌŒ «·ÿ»«⁄…|" & " «—ÌŒ «· ”·Ì„|"
    .ColWidth(0) = 1000
    .ColWidth(1) = 4500
    .ColWidth(2) = 2000
    .ColWidth(3) = 1350
    .ColWidth(4) = 1350
    
    .ColHidden(0) = True
    .ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 2
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub Fixgrd3()
With grid3
    .FormatString = "—ﬁ„ «·„ÿ«·»…|" & "—ﬁ„ «·«” „«—…|" & " «—ÌŒ «·«” „«—…|" & "«·„Ê”„|" & "‰Ê⁄ «·„ÿ«·»…|" & " «—ÌŒ «·«‰ Â«¡|" & "«‘ —«ﬂ«  «·”‰…|" & "«‘ —«ﬂ«  „ √Œ—…|" & "ﬁÌ„… „÷«›…|" & "€—«„…  √ŒÌ—|" & "≈Ã„«·Ì «·«” „«—…|"
    .ColWidth(0) = 1300
    .ColWidth(1) = 1300
    .ColWidth(2) = 1400
    .ColWidth(3) = 1500
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
    .ColWidth(8) = 1400
    
    .ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 2
        .ColAlignment(I) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub cellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid2
KeyCode = 0
If Col < .Cols - 2 Then
    .Col = 4
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid2, Row + 1, 4, 4)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem2()
With grid2
.AddItem ""
'.TextMatrix(.rows - 1, 0) = myFormat_p(Date)
End With
End Sub
Private Function myreplaceGrd2(Row) As Boolean
Dim aInsert As Variant
With grid2
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid2.rows - 2, Row)
        aInsert = AddFlag(Empty, "[DATE_DELIVERY]", addDate(grid2.TextMatrix(I, 4)))
'        If grid2.TextMatrix(i, grid2.Cols - 1) = "" Then
'            con.Execute addInsert(aInsert, "FILE4_10")
'        Else
            con.Execute addUpdate(aInsert, "FILE4_10", "ID = " & grid2.TextMatrix(I, .Cols - 1))
'        End If
    Next
End With
myreplaceGrd2 = True
End Function
Private Sub myloadgrd2()
With grid2
Dim cString As String
cString = "SELECT CODE,DESCA,CASE WHEN CODE IS NULL THEN '«·⁄÷Ê ‰›”Â' ELSE  RELATION_DESCA END, CONVERT(VARCHAR(10),DATE,111),CONVERT(VARCHAR(10), DATE_DELIVERY,111),FILE4_10.ID" & _
           " From FILE4_10"
cString = cString & " WHERE FILE4_10.MEMBER = " & xCode.text
cString = cString & " AND FILE4_10.YEAR = " & sSeason
cString = cString & " ORDER BY FILE4_10.MEMBER "
Set DATA12.Recordset = myCmd(cString, con)
'myAddItem2
fixgrd2
End With
End Sub
Private Sub Myloadgrd3()
With grid3
Dim cString As String
cString = "SELECT FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,CONVERT(VARCHAR(10), FILE6_20H.DATE,111), YEARS_CODES.DESCA_R, PAID_TYPES.DESCA, CONVERT(VARCHAR(10),YEARS_CODES.DATE2,111), FILE6_20H.TOTAL_YEAR" & _
        ",FILE6_20H.TOTAL_YEAR_OTHER , FILE6_20H.TOTAL_TAX, FILE6_20H.TOTAL_LATE, FILE6_20H.TOTAL" & _
        " FROM  FILE6_20H INNER JOIN YEARS_CODES ON FILE6_20H.YEAR_CODE = YEARS_CODES.CODE INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE"
cString = cString & " WHERE FILE6_20H.CODE = " & addvalue(xCode.text)
cString = cString & " ORDER BY FILE6_20H.DATE DESC"
Set DATA13.Recordset = myRecordSet(cString, con)
Fixgrd3
If grid3.rows > 1 Then
    grid3.ShowCell 1, 0
    grid3.Select 1, 0
End If
End With
End Sub

Private Sub myRemove2(Row As Long)
grid2.RemoveItem Row
End Sub
Private Function LoadPhoto(pCode As String) As Boolean
On Error Resume Next
xMemberPhoto.Picture = LoadPicture("")
If Dir(retPhoto(pCode)) <> "" Then xMemberPhoto.Picture = LoadPicture(retPhoto(pCode))
Err.Clear
End Function
Private Function loadPhoto_Append(pCode As String, Optional pAppend As String = "") As Boolean
On Error Resume Next
xAppendPhoto.Picture = LoadPicture("")
If Dir(RetAppendPhoto(pCode, pAppend)) <> "" Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(pCode, pAppend))
Err.Clear
End Function
Private Sub GetMembers()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
con.Execute "delete from file1_10"

Set loctable = New ADODB.Recordset
loctable.Open "select * from members where membership_type = 1", con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount
Do Until loctable.EOF
    I = I + 1
    Me.Caption = I & " from " & nRecords
    aInsert = AddFlag(Empty, "ID", addvalue(loctable!ID))
    aInsert = AddFlag(aInsert, "CODE", addvalue(loctable!NO))
    aInsert = AddFlag(aInsert, "TITLE", addstring(loctable!Title))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[DATE_BIRTH]", addDate(loctable!birthdate))
    aInsert = AddFlag(aInsert, "[DATE_BEGIN]", addDate(loctable!ORDER_DATE))
    aInsert = AddFlag(aInsert, "[SES_NO]", addstring(loctable!ORDER_))
    
    aInsert = AddFlag(aInsert, "[ID_NO]", addstring(loctable!ID_NO))
    aInsert = AddFlag(aInsert, "[religion]", addvalue(loctable!RELIGION))
    aInsert = AddFlag(aInsert, "[REGION]", addvalue(loctable!Area))
    aInsert = AddFlag(aInsert, "[gender]", addvalue(loctable!GENDER))
    aInsert = AddFlag(aInsert, "[DEGREE]", addvalue(loctable!QUAL))
    aInsert = AddFlag(aInsert, "[SOCIAL]", addvalue(loctable!marital_status))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!home_address))
    aInsert = AddFlag(aInsert, "[PHONE]", addstring(loctable!HOME_PHONE))
    aInsert = AddFlag(aInsert, "[MOBIL]", addstring(loctable!MOBILE))
    aInsert = AddFlag(aInsert, "[JOB]", addvalue(loctable!job))
    aInsert = AddFlag(aInsert, "[JOB_DESCA]", addstring(loctable!work_add))
    aInsert = AddFlag(aInsert, "[PHONE_WORK]", addstring(loctable!WORK_PHONE))
    aInsert = AddFlag(aInsert, "[TYPE]", addvalue(loctable!mem_section))
    aInsert = AddFlag(aInsert, "[DATE_LAST]", addDate(loctable!RENEW_DATE))
    aInsert = AddFlag(aInsert, "[DATE_END]", addDate(loctable!END_DATE))
    aInsert = AddFlag(aInsert, "[ORDER_DATE]", addDate(loctable!ORDER_DATE))
    aInsert = AddFlag(aInsert, "[CARD_NO]", addstring(loctable!CARD_NO))
    aInsert = AddFlag(aInsert, "[RELCARD]", addstring(loctable!RELCARD))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!notes))
    aInsert = AddFlag(aInsert, "[DEAD]", IIf(mRound(loctable!DEAD), "1", "0"))
    aInsert = AddFlag(aInsert, "[DROP]", IIf(mRound(loctable!Drop), "1", "0"))
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!MDATE))
    aInsert = AddFlag(aInsert, "[membership_type]", addvalue(loctable!membership_type))
    aInsert = AddFlag(aInsert, "[SUBVAL]", mRound(loctable!SUBVAL))
    con.Execute addInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
MsgBox "DONE MEMBER"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub GetRelation()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
con.Execute "delete from file1_11"

Set loctable = New ADODB.Recordset
loctable.Open "select relatives.* from relatives inner join members on relatives.member = members.no where membership_type = 1 ", con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount
Do Until loctable.EOF
    I = I + 1
    Me.Caption = I & " from " & nRecords
    aInsert = AddFlag(Empty, "CODE", addvalue(loctable!rel_order))
    aInsert = AddFlag(aInsert, "MEMBER", addstring(loctable!member))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!Name))
    If mRound(loctable!RELATION) = 1 Then
        aInsert = AddFlag(aInsert, "[RELATION]", addvalue(loctable!RELATION))
    Else
        aInsert = AddFlag(aInsert, "[RELATION]", mRound(loctable!RELATION) - 1)
    End If
    aInsert = AddFlag(aInsert, "[GENDER]", addvalue(loctable!GENDER))
    aInsert = AddFlag(aInsert, "[DATE_BIRTH]", addDate(loctable!birth_date))
    aInsert = AddFlag(aInsert, "[DATE_BEGIN]", addDate(loctable!BEGIN_DATE))
    aInsert = AddFlag(aInsert, "[CARD]", addstring(loctable!card))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!REMARKS))
    aInsert = AddFlag(aInsert, "[MDATE]", addDate(loctable!MDATE))
    If IsNull(loctable!Paid) Then
        aInsert = AddFlag(aInsert, "[PAID]", "0")
    Else
        aInsert = AddFlag(aInsert, "[PAID]", IIf(loctable!Paid, 1, 0))
    End If
    con.Execute addInsert(aInsert, "FILE1_11")
    loctable.MoveNext
Loop
MsgBox "DONE MEMBER"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub GetPaidItems()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
'On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
con.Execute "delete from file6_11"

Set loctable = New ADODB.Recordset
loctable.Open "select [claim_values].* from [claim_values]", con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount


aTypes = GetRows("select * from type_codes", con)
aclaim = GetRows("select * from Paid_Types", con)
For n = 0 To UBound(aclaim)
    For I = 0 To UBound(aTypes)
        loctable.MoveFirst
        nRow = 0
        Do Until loctable.EOF
            nRow = nRow + 1
            Me.Caption = nRow & " from " & nRecords
            aInsert = AddFlag(Empty, "[TYPE]", addvalue(retFlag(aclaim(n), "CODE")))
            aInsert = AddFlag(aInsert, "[SECTION]", addvalue(retFlag(aTypes(I), "CODE")))
            aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
            aInsert = AddFlag(aInsert, "VALUE", mRound(loctable!Value))
            aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(loctable!Year))
            con.Execute addInsert(aInsert, "FILE6_11")
            loctable.MoveNext
        Loop
    Next
Next
MsgBox "DONE PAYMENITEMS"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub GetPaidItems2()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
'On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
con.Execute "delete from file6_11 Where YEAR_CODE < 22"

Set loctable = New ADODB.Recordset

cString = "SELECT claim_groups.type, claim_groups.item, claim_groups.season, claim_values.value" & _
          " FROM claim_groups INNER JOIN   claim_values ON claim_groups.item = claim_values.item AND claim_groups.season = claim_values.season"
loctable.Open cString, con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount
aSection = GetRows("select * from type_codes", con)
For I = 0 To UBound(aSection)
    loctable.MoveFirst
    nRow = 0
    Do Until loctable.EOF
        nRow = nRow + 1
        Me.Caption = nRow & " from " & nRecords
        aInsert = AddFlag(Empty, "[TYPE]", loctable!Type)
        aInsert = AddFlag(aInsert, "[SECTION]", addvalue(retFlag(aSection(I), "CODE")))
        aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
        aInsert = AddFlag(aInsert, "VALUE", mRound(loctable!Value))
        aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(loctable!SEASON))
        aInsert = AddFlag(aInsert, "[BASIC]", "1")
        con.Execute addInsert(aInsert, "FILE6_11")
        loctable.MoveNext
    Loop
Next

Set loctable = Nothing
Set loctable = New ADODB.Recordset

cString = "SELECT [claim_item],[mem_section],[discount],[season]  From [Olympic2].[dbo].[discounts]"
loctable.Open cString, con2, adOpenStatic, adLockReadOnly, adCmdText

nRow = 0
nRecords = loctable.RecordCount
Do Until loctable.EOF
    nRow = nRow + 1
    Me.Caption = nRow & " from " & nRecords
    cWhere = "FILE6_11.ITEM = " & loctable!claim_item
    cWhere = cWhere & " AND " & "FILE6_11.[SECTION] = " & loctable!mem_section
    cWhere = cWhere & " AND " & "FILE6_11.YEAR_CODE = " & loctable!SEASON
    con.Execute "update FILE6_11 SET FILE6_11.DISCOUNT = " & mRound(loctable!discount) & " WHERE " & cWhere
    loctable.MoveNext
Loop

con.Execute "UPDATE FILE6_11 SET FILE6_11.YEAR_CODE = YEARS_CODES.CODE FROM FILE6_11 INNER JOIN YEARS_CODES ON FILE6_11.year_code = YEARS_CODES.CODE_SYSTEM"

MsgBox "DONE PAYMENITEMS"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub GetDocHeader()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
con.Execute "delete from file6_20h"
If False Then
Set loctable = New ADODB.Recordset
loctable.Open "select claims.* from claims inner join members on claims.member = members.no where membership_type = 1", con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount
Do Until loctable.EOF
    I = I + 1
    Me.Caption = I & " from " & nRecords
    aInsert = AddFlag(Empty, "DOC_NO", addvalue(loctable!ID))
    aInsert = AddFlag(aInsert, "SEASON", addvalue(loctable!SEASON))
    aInsert = AddFlag(aInsert, "CODE", addvalue(loctable!member))
    aInsert = AddFlag(aInsert, "Form_no", mRound(loctable!doc_paper))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!issue_date))
    aInsert = AddFlag(aInsert, "[TYPE]", addvalue(loctable!claim_type))
    aInsert = AddFlag(aInsert, "[YEAR_CODE]", addvalue(loctable!claim_year))
    aInsert = AddFlag(aInsert, "[LATE_VALUE]", mRound(loctable!LATE_VALUE))
    aInsert = AddFlag(aInsert, "[TOTAL_VALUE]", mRound(loctable!claim_value))
    con.Execute addInsert(aInsert, "FILE6_20H")
    loctable.MoveNext
Loop
End If
con.Execute "UPDATE FILE6_20H SET FILE6_20H.YEAR_CODE = YEARS_CODES.CODE FROM FILE6_20H INNER JOIN  YEARS_CODES ON FILE6_20H.SEASON = YEARS_CODES.CODE_SYSTEM"
MsgBox "DONE MEMBER"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub GetFooter()
Dim aInsert As Variant
Dim loctable As ADODB.Recordset
'On Error GoTo myerror
Dim con2 As New ADODB.Connection
openCon con2, LoadConString(, "Olympic2")
    
'con.Execute "delete from file6_20"

Set loctable = New ADODB.Recordset
loctable.Open "select  claim_details.* from claim_details inner join claims on claim_details.doc = claims.id inner join members on claims.member = members.no where membership_type = 1 order by doc,item,year,season", con2, adOpenStatic, adLockReadOnly, adCmdText
nRecords = loctable.RecordCount
Do Until loctable.EOF
    I = I + 1
    If I > 3200107 Then
        Me.Caption = I & " from " & nRecords
        nRecord = nRecord + 1
        aInsert = AddFlag(Empty, "DOC_NO", addvalue(loctable!Doc))
        aInsert = AddFlag(aInsert, "SEASON", addvalue(loctable!SEASON))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
        aInsert = AddFlag(aInsert, "VALUE", mRound(loctable!Value))
        aInsert = AddFlag(aInsert, "[QUANT]", mRound(loctable!Number))
        aInsert = AddFlag(aInsert, "[DISCOUNT_RATE]", mRound(loctable!discount))
        cInsert = addInsert(aInsert, "FILE6_20") & ";"
        con.Execute cInsert
    End If
    loctable.MoveNext
Loop


'For i = 0 To UBound(asql)
'    Me.Caption = i + 1 & " from " & (UBound(asql) + 1)
'    con.Execute asql(i)
'Next

MsgBox "DONE MEMBER"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub FixTotals()
Dim cString As String
Dim loctable As New ADODB.Recordset
loctable.Open "select code from years_codes order by code", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = nRecord
    cString = "UPDATE FILE6_20H SET FILE6_20H.TOTAL_YEAR = dbo.f_inv_total_year(FILE6_20H.DOC_NO)," & _
              " FILE6_20H.TOTAL_YEAR_OTHER = dbo.f_inv_total_year_other(FILE6_20H.DOC_NO),FILE6_20H.TOTAL_LATE = dbo.f_inv_total_late(FILE6_20H.DOC_NO)," & _
                " FILE6_20H.TOTAL_TAX = dbo.f_inv_total_tax(FILE6_20H.DOC_NO)" & _
               " FROM FILE6_20H WHERE YEAR_CODE = " & loctable!code
    con.Execute cString
    loctable.MoveNext
Loop
End Sub
Private Sub HandleFirst()
cmdClaim.Enabled = sCode = ""
End Sub
Public Sub myrefresh()
myUndo
End Sub
Private Function SetMemberCard(pCode, pCard) As Integer
On Error GoTo myerror
SetMemberCard = -1
sFound = GetField("Select code from file1_10 where card = " & MyParn(pCard) & " and code <> " & pCode) & ""
If sFound <> "" Then
    MsgBox "—ﬁ„ «·ﬂ«—  „ÊÃÊœ ›Ì ⁄÷ÊÌ… —ﬁ„ " & sFound
    Exit Function
End If
con.Execute "update file1_10 set " & _
            " file1_10.card = " & MyParn(pCard) & _
            " where file1_10.code = " & addvalue(pCode), SetMemberCard
If SetMemberCard = 0 Then MsgBox "·„ Ì „ Õ›Ÿ »Ì«‰«  «·ﬂ«— "
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function SetRelCard(pCode, pRel, pCard) As Integer
On Error GoTo myerror
Dim sFound As String
SetRelCard = -1
sFound = GetField("Select Member from file1_11 where card = " & MyParn(pCard) & " and not (member = " & pCode & " and Code = " & pRel & ")") & ""
If sFound <> "" Then
    MsgBox "—ﬁ„ «·ﬂ«—  „ÊÃÊœ ›Ì ⁄÷ÊÌ… —ﬁ„ " & sFound
    Exit Function
End If
con.Execute "update file1_11 set " & _
            " file1_11.card = " & MyParn(pCard) & _
            " where file1_11.Member = " & addvalue(pCode) & _
            " and file1_11.Code = " & addvalue(pRel) _
            , SetRelCard
If SetRelCard = 0 Then MsgBox "·„ Ì „ Õ›Ÿ »Ì«‰«  «·ﬂ«— "
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub AddOtherType()
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select top 1 code,desca from paid_types where type = 200", con)
If Not loctable.EOF Then
    cmdService.Caption = "⁄—÷ »‰Êœ „ÿ«·»… " & loctable!Desca
    cmdService.Tag = loctable!code
    cmdService.TagVariant = loctable!Desca & ""
End If
End Sub

