VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form member_invfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «⁄÷«¡ «·œ⁄Ê…"
   ClientHeight    =   9930
   ClientLeft      =   615
   ClientTop       =   1320
   ClientWidth     =   15615
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
   ScaleHeight     =   9930
   ScaleWidth      =   15615
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   0
      Width           =   2310
      Begin VB.TextBox xCode_Trans 
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
         Height          =   435
         Left            =   45
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   180
         Width           =   1500
      End
      Begin VB.CommandButton cmdTrans 
         Caption         =   "‰ﬁ·"
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
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   88
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   690
      End
   End
   Begin VB.Frame Frame11 
      Height          =   2130
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   3825
      Width           =   4110
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
         Height          =   1830
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   3930
      End
      Begin VB.Label Label17 
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
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   585
         Width           =   990
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   8955
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   180
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
         Picture         =   "member_inv.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_inv.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   855
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   135
         Width           =   735
         _ExtentX        =   1296
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
         Picture         =   "member_inv.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_inv.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1620
         TabIndex        =   84
         TabStop         =   0   'False
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
         Picture         =   "member_inv.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_inv.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2385
         TabIndex        =   85
         TabStop         =   0   'False
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
         Picture         =   "member_inv.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "member_inv.frx":EB26
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1005
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   4950
      Width           =   10950
      Begin VB.TextBox xMonths 
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
         Left            =   3105
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Tag             =   "D"
         Top             =   540
         Width           =   465
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
         Left            =   4545
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Tag             =   "D"
         Top             =   540
         Width           =   1365
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
         Left            =   3105
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "D"
         Top             =   180
         Width           =   2805
      End
      Begin MSDataListLib.DataCombo xReason 
         Height          =   330
         Left            =   7425
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
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   7425
         TabIndex        =   23
         Top             =   540
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
      Begin MSDataListLib.DataCombo xPaid_code 
         Height          =   330
         Left            =   45
         TabIndex        =   27
         Top             =   135
         Width           =   1725
         _ExtentX        =   3043
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
      Begin VB.Label Label8 
         Caption         =   "‰Ê⁄ «·”œ«œ"
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
         TabIndex        =   91
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "⁄œœ «·‘ÂÊ—"
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label25 
         Caption         =   " «—ÌŒ «·⁄÷ÊÌ…"
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   585
         Width           =   1125
      End
      Begin VB.Label Label28 
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label15 
         Caption         =   "—ﬁ„ «·Ã·”…"
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "«”«” «·⁄÷ÊÌ…"
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   225
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   3240
      Width           =   4110
      Begin VB.Label Label20 
         Caption         =   " «—ÌŒ «Œ— ÿ»«⁄…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label xdate_Print 
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
         TabIndex        =   74
         Top             =   180
         Width           =   2490
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1545
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   9855
      Visible         =   0   'False
      Width           =   5550
      Begin VB.CommandButton Command2 
         Caption         =   "«÷«›… «·«⁄÷«¡"
         Height          =   600
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   765
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «·’Ê—"
         Height          =   600
         Left            =   -405
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1485
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… «· Ê«»⁄"
         Height          =   600
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         Caption         =   "«÷«›… ’Ê— «· Ê«»⁄"
         Height          =   600
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   495
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command5"
         Height          =   420
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   405
         Visible         =   0   'False
         Width           =   3075
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1275
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2745
      Width           =   10995
      Begin VB.TextBox xFace 
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
         Left            =   90
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   855
         Width           =   3930
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
         Left            =   5400
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   855
         Width           =   4110
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
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   135
         Width           =   9420
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
         Left            =   5400
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   495
         Width           =   4110
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
         Left            =   90
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   495
         Width           =   3930
      End
      Begin VB.Label Label23 
         Caption         =   "Face Book"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   900
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   855
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   180
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
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   540
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   540
         Width           =   1260
      End
   End
   Begin VB.Frame Frame7 
      Height          =   690
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   0
      Width           =   8700
      Begin VB.CommandButton cmdInform_rel 
         Caption         =   "«” ⁄·«„  «»⁄"
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
         Left            =   6210
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
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
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_inv.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2505
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_inv.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "member_inv.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_inv.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4965
         MaskColor       =   &H00FFFFFF&
         Picture         =   "member_inv.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   7425
         Picture         =   "member_inv.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   675
      Width           =   10995
      Begin VB.CommandButton cmdRegion 
         Caption         =   "..."
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   945
         Width           =   375
      End
      Begin VB.CommandButton cmdDegree 
         Caption         =   "..."
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
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
         Left            =   5625
         MaxLength       =   14
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1620
         Width           =   3975
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
         Left            =   5625
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   1260
         Width           =   3975
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
         Left            =   5625
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   3975
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
         Left            =   5625
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   3975
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
         Left            =   8145
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xGender 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   1305
         Width           =   3390
         _ExtentX        =   5980
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
         Left            =   135
         TabIndex        =   11
         Top             =   1665
         Width           =   3390
         _ExtentX        =   5980
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
         TabIndex        =   5
         Top             =   225
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
      Begin MSDataListLib.DataCombo xDegree 
         Height          =   330
         Left            =   540
         TabIndex        =   6
         Top             =   585
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
      Begin MSDataListLib.DataCombo xRegion 
         Height          =   330
         Left            =   540
         TabIndex        =   8
         Top             =   945
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
      Begin VB.Label Label27 
         Caption         =   "«· ﬁ”Ì„ «·«œ«—Ì"
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   990
         Width           =   1215
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
         Index           =   1
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1620
         Width           =   1035
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   630
         Width           =   855
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   1335
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1710
         Width           =   765
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
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1350
         Width           =   495
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1350
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label7 
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label2 
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   630
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3420
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   -1125
      Top             =   1260
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
      Left            =   2520
      Top             =   7020
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
      Top             =   495
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
      Left            =   -1215
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
      Left            =   -945
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
      Left            =   -585
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   375
      Left            =   -450
      Top             =   720
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
      TabIndex        =   67
      Top             =   9585
      Width           =   15615
      _ExtentX        =   27543
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
      Left            =   -630
      Top             =   855
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
      Height          =   3210
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   45
      Width           =   4110
      Begin Threed.SSCommand cmdScan 
         Height          =   555
         Left            =   90
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2565
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
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
         Picture         =   "member_inv.frx":1EFD6
         Caption         =   "„”Õ ÷Ê∆Ì"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "member_inv.frx":21770
      End
      Begin VB.Image xAppendPhoto 
         Appearance      =   0  'Flat
         Height          =   2310
         Left            =   135
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1950
      End
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         Height          =   2310
         Left            =   2115
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1860
      End
   End
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "member_inv.frx":244CB
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   -270
      TabIndex        =   47
      Tag             =   "-1"
      Top             =   270
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrmind71"
      LicenseRegCode  =   "íß“Ωª∫≠Ω≥´±“™ºØ´¥æÆØUBOR-FEOEONZI-EPCP6gI"
   End
   Begin VB.Frame Frame6 
      Height          =   960
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4005
      Width           =   10995
      Begin VB.CommandButton cmdCompany 
         Caption         =   "..."
         Height          =   330
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
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
         Height          =   345
         Left            =   90
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "D"
         Top             =   180
         Width           =   3975
      End
      Begin VB.CommandButton cmdJob 
         Caption         =   "..."
         Height          =   330
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   375
      End
      Begin MSDataListLib.DataCombo xJob 
         Height          =   330
         Left            =   6210
         TabIndex        =   17
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
         Left            =   6210
         TabIndex        =   19
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
      Begin VB.Label Label26 
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label Label10 
         Caption         =   "»Ì«‰ «·⁄„·"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   225
         Width           =   1125
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   225
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc data9 
      Height          =   420
      Left            =   3420
      Top             =   -270
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
      Height          =   2940
      Left            =   45
      TabIndex        =   33
      Top             =   5985
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   5186
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Ã“«¡«  «··«⁄»Ì‰"
      TabPicture(0)   =   "member_inv.frx":2490D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "»Ì«‰«  «· Ê«»⁄"
      TabPicture(1)   =   "member_inv.frx":24929
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   2490
         Left            =   90
         TabIndex        =   29
         Top             =   360
         Width           =   14955
         _cx             =   26379
         _cy             =   4392
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
         BackColorFixed  =   12648384
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   2445
         Left            =   -74910
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   360
         Width           =   14955
         _cx             =   26379
         _cy             =   4313
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
         BackColorFixed  =   12648384
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483640
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
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
   End
   Begin MSAdodcLib.Adodc DATA12 
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
      Height          =   465
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   9090
      Width           =   4020
   End
End
Attribute VB_Name = "member_invfrm"
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
Dim oSearch As New Search, oSearchRel As New Search
Dim CardTable As ADODB.Recordset
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bEdit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
cmddel.Enabled = (nMode = LoadMode And bEditRecord)
cmdSave.Enabled = bEditRecord
cmdInform.Enabled = (nMode = LoadMode)

aRecords = retRecords(xCode.Text)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    xRecordNo.Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    xRecordNo.Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sCode = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sCode = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sCode = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sCode = ""

xCode.Enabled = bEdit = Not (nMode = LoadMode)
cmdScan.Enabled = nMode = LoadMode And bEditRecord
'cmdScan2.Enabled = nMode = LoadMode And bEditRecord
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = Newflag("FILE1_50", "code")
xTitle.Text = ""
xdesca.Text = ""
xJob_desca.Text = ""
xDate_birth.Text = ""
xCompany.BoundText = ""
xGender.BoundText = "1"
xSocial.BoundText = ""
xReason.BoundText = ""
xPaid_code.Text = ""
xMonths.Text = ""

xDate_Begin.Text = ""
xSes_no.Text = ""

xFace.Text = ""
xReligion.BoundText = "1"
xId_no.Text = ""
xAddress.Text = ""
xPhone.Text = ""
xMobil.Text = ""
xMail.Text = ""
xJob.BoundText = ""
xDegree.BoundText = ""
xRegion.BoundText = ""
xType.BoundText = ""
'xDate_Last.Text = ""
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

xdate_Print.Caption = ""
'xdate_paid.Caption = ""
StatusBar1.Panels(1).Text = ""
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(4).Text = ""

Fixgrd
grid1.rows = 1
myAddItem

Fixgrd
grid2.rows = 1
myAddItem2

Handlecontrols DefineMode
xRecordNo.Caption = "«÷«›… ”Ã· ÃœÌœ " & "(" & CardTable.RecordCount & ")"
On Error Resume Next
CellPos 13, grid1.rows - 2, grid1.Cols - 1
grid1.SetFocus
Err.Clear
If SSTab1.Tab = 0 Then SSTab1.Tab = 1
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
    myUndo
ElseIf ActiveControl.Name = Me.cmdInform_rel.Name Then
    xCode.Text = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 0)
    oSearchRel.Hide
    myUndo
End If
End Sub
Private Sub myload()
xCode.Text = CardTable!CODE & ""
xTitle.Text = CardTable!Title & ""
xdesca.Text = CardTable!Desca & ""
xFace.Text = CardTable!Face & ""
'xFace.Text = CardTable!Face & ""
xDate_birth.Text = myFormat_p(CardTable!DATE_BIRTH)
xSes_no.Text = CardTable!SES_NO & ""

xGender.BoundText = CardTable!GENDER & ""
xSocial.BoundText = CardTable!SOCIAL & ""
xCompany.BoundText = CardTable!COMPANY & ""
xReligion.BoundText = CardTable!RELIGION & ""
xId_no.Text = CardTable!ID_NO & ""
xAddress.Text = CardTable!Address & ""
xJob_desca.Text = CardTable!JOB_DESCA & ""
xPhone.Text = CardTable!Phone & ""
xMobil.Text = CardTable!MOBIL & ""
xMail.Text = CardTable!MAIL & ""
xJob.BoundText = CardTable!job & ""
xJob_desca.Text = CardTable!JOB_DESCA & ""
xDegree.BoundText = CardTable!Degree & ""
xCompany.BoundText = CardTable!COMPANY & ""
xRegion.BoundText = CardTable!REGION & ""
xReason.BoundText = CardTable!REASON & ""
xMonths.Text = Myvalue(CardTable!Months)
xPaid_code.BoundText = CardTable!PAID_CODE & ""

xDate_Begin.Text = myFormat_p(CardTable!DATE_BEGIN)
xType.BoundText = CardTable!Type & ""
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

StatusBar1.Panels(1).Text = CardTable!UserName & ""
StatusBar1.Panels(2).Text = myFormat_p(CardTable!Time, True)
StatusBar1.Panels(3).Text = CardTable!UserName2 & ""
StatusBar1.Panels(4).Text = myFormat_p(CardTable!Time2, True)

'xRecordNo.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount

If validPhoto(RetPhoto_I(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto_I(xCode.Text))
'aret = LastDoc(xCode.Text, con)
'xDoc_No.Caption = retFlag(aret, "FORM_NO") & ""
'xdate_paid.Caption = myFormat_p(CardTable!date_paid)
xdate_Print.Caption = myFormat_p(CardTable!DATE_PRINT)
myloadGrd
myloadgrd2
CellPos 13, 0, grid1.Cols - 1
cellPos2 13, 0, grid2.Cols - 1
On Error Resume Next
grid1.SetFocus
loadPhoto_Append xCode.Text, grid1.TextMatrix(grid1.Row, 0)
Err.Clear
'If SSTab1.Tab = 0 Then grid1.SetFocus Else grid2.SetFocus
End Sub
Private Function MyReplace(Optional Row As Long = -1, Optional Row2 As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "Title", addstring(xTitle.Text))
aInsert = AddFlag(aInsert, "Desca", addstring(xdesca.Text))
aInsert = AddFlag(aInsert, "Date_birth", addstring(xDate_birth.Text))
aInsert = AddFlag(aInsert, "Date_Begin", addDate(xDate_Begin.Text))

aInsert = AddFlag(aInsert, "SES_NO", addstring(xSes_no.Text))
aInsert = AddFlag(aInsert, "Gender", addvalue(xGender.BoundText))
aInsert = AddFlag(aInsert, "Social", addvalue(xSocial.BoundText))
aInsert = AddFlag(aInsert, "Religion", addvalue(xReligion.BoundText))
aInsert = AddFlag(aInsert, "Id_no", addstring(xId_no.Text))
aInsert = AddFlag(aInsert, "Address", addstring(xAddress.Text))
aInsert = AddFlag(aInsert, "Phone", addstring(xPhone.Text))
aInsert = AddFlag(aInsert, "Mobil", addstring(xMobil.Text))
aInsert = AddFlag(aInsert, "Mail", addstring(xMail.Text))
aInsert = AddFlag(aInsert, "Job", addvalue(xJob.BoundText))
aInsert = AddFlag(aInsert, "FACE", addstring(xFace.Text))
aInsert = AddFlag(aInsert, "Degree", addvalue(xDegree.BoundText))
aInsert = AddFlag(aInsert, "Region", addvalue(xRegion.BoundText))
aInsert = AddFlag(aInsert, "company", addvalue(xCompany.BoundText))
aInsert = AddFlag(aInsert, "Job_desca", addstring(xJob_desca.Text))
aInsert = AddFlag(aInsert, "Type", addvalue(xType.BoundText))
aInsert = AddFlag(aInsert, "reason", addvalue(xReason.BoundText))
aInsert = AddFlag(aInsert, "months", Val(xMonths.Text))
aInsert = AddFlag(aInsert, "PAID_CODE", addvalue(xPaid_code.BoundText))

con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    aInsert = AddFlag(aInsert, "Code", addvalue(xCode.Text))
    con.Execute addInsert(aInsert, "FILE1_50")
Else
    con.Execute addUpdate(aInsert, "FILE1_50", "FILE1_50.CODE = " & addvalue(xCode.Text))
End If
If (Row = -1 And Row2 = -1) Or Row <> -1 Then myreplaceGrd Row
If (Row = -1 And Row2 = -1) Or Row2 <> -1 Then myreplaceGrd2 Row2
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub

Private Sub cmdCompany_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xCompany.BoundText
oFlagfrm.sTable = "company_CODES"
oFlagfrm.sCaption = "«·‘—ﬂ…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
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
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data6.Recordset.Requery
xDegree.BoundText = sBound
If Not xDegree.MatchedWithList Then xDegree.BoundText = ""
End Sub

Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If grid1.rows > 2 Then
        MsgBox "«·⁄÷Ê ·Â  Ê«»⁄ ÌÃ» Õ–›Â„ «Ê·«"
        Exit Sub
    End If
    con.BeginTrans
    con.Execute "Delete  From FILE1_50 Where code = " & xCode.Text
    
    DeletePhoto_I xCode.Text
    con.CommitTrans
    
    openCardTable xCode.Text, "<="
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
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
If Trim(xCode.Text) = "" Then Exit Sub
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
        fs.CopyFile cFile, RetPhoto_I(xCode.Text)
    End If
    myload
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub CmdNext_Click()
openCardTable xCode.Text, ">"
If CardTable.EOF Then openCardTable xCode.Text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.Text, "<"
If CardTable.EOF Then openCardTable xCode.Text, "="
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
MemberLookupAll_I Me, oSearch, cFilter
End Sub
Private Sub cmdInform_rel_Click()
relLookupAll_I Me, oSearchRel
End Sub

Private Sub cmdJob_Click()
Dim oFlagfrm As New flag_mainfrm, sBound As String
sBound = xJob.BoundText
oFlagfrm.sTable = "JOB_CODES"
oFlagfrm.sCaption = "«·ÊŸÌ›…"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
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
oFlagfrm.bEdit = True
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
FlagFrm.bEdit = True
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
oFlagfrm.bEdit = True
oFlagfrm.Show 1
DATA7.Recordset.Requery
xRegion.BoundText = sBound
If Not xRegion.MatchedWithList Then xRegion.BoundText = ""
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
'openCardTable
myUndo
End Sub
Private Sub cmdScan_Click()
Scan_ifrm.sCode = xCode.Text
Scan_ifrm.Show 1
If validPhoto(RetPhoto_I(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto_I(xCode.Text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto_i(xCode.Text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto_i(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
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

Private Sub cmdTrans_Click()
If Not validTrans Then Exit Sub
If myTrans Then Inform " „ ‰ﬁ· »Ì«‰«  «·⁄÷Ê »‰Ã«Õ"
'createAddInsertTable "SELECT * FROM file1_11", con
End Sub

Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub
Private Sub cmdScan2_Click()
nPhoto = 0
ScanImage
On Error Resume Next
If validPhoto(RetPhoto_I(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto_I(xCode.Text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto_i(xCode.Text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto_i(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
End If
Err.Clear
End Sub
Private Sub Command1_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, i As Long, cFile As String, nRecordcount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
'loctable.Open "select * from FILE1_50 where NEWDATA = true", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.Open "select * from FILE1_50", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordcount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    i = i + 1
    Me.Caption = cCaption & i & " from " & nRecordcount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, RetPhoto_I(loctable!CODE)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub

Private Sub Command2_Click()
AddMember
End Sub
Private Sub AddMember()
Dim conMdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror
conMdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_50", conMdb, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(Empty, "CODE", loctable!CODE)
    aInsert = AddFlag(aInsert, "MEMBERID", addstring(loctable!MEMBERID))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca & ""))
    aInsert = AddFlag(aInsert, "SECTION", addvalue(loctable!Section & ""))
    aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(Format(loctable!DATE_BIRTH, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "union_reg", addstring(loctable!Union_reg & ""))
    aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!notes & ""))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address & ""))
    aInsert = AddFlag(aInsert, "JOB_CODE", addstring(loctable!JOB_CODE & ""))
    aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!Phone & ""))
    aInsert = AddFlag(aInsert, "MOBIL", addstring(loctable!MOBIL & ""))
    aInsert = AddFlag(aInsert, "PHOTO_CODE", addstring(loctable!PHOTO_CODE & ""))
    con.Execute addInsert(aInsert, "FILE1_50")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conMdb.Close
Set conMdb = Nothing
Exit Sub
myerror:
MsgBox Err.Description
End Sub
Private Sub addRelation()
Dim conMdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
'On Error GoTo myerror
conMdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_51", conMdb, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(aInsert, "MEMBER", addvalue(loctable!Member))
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
    con.Execute addInsert(aInsert, "FILE1_51")
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

Private Sub Command3_Click()
addRelation
End Sub

Private Sub Command4_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
'Set f = fs.GetFolder(App.Path & "\photo\")
'Set fc = f.Files
'nCount = fc.Count
Dim cString As String, i As Long, cFile As String, nRecordcount As Long, cCaption As String
Dim loctable As New ADODB.Recordset
loctable.Open "select * from FILE1_51", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.MoveLast
nRecordcount = loctable.RecordCount
loctable.MoveFirst
cCaption = Me.Caption
Do Until loctable.EOF
    i = i + 1
    Me.Caption = cCaption & i & " from " & nRecordcount
    If Not IsNull(loctable!PHOTO_CODE) Then
        cFile = App.Path & "\person\" & loctable!PHOTO_CODE
        If fs.FileExists(cFile) Then
            fs.CopyFile cFile, RetAppendPhoto_i(loctable!Member, loctable!CODE)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub

Private Sub Command5_Click()
'createAddInsert "SELECT * FROM FILE1_50", con
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If xCode.Tag = LoadMode Then
        If SSTab1.Tab = 1 Then
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

Set data8.Recordset = myRecordSet("select * from reason_Codes", con)
Set xReason.RowSource = data8
xReason.ListField = "Desca"
xReason.BoundColumn = "Code"

Set data9.Recordset = myRecordSet("select * from type_Codes", con)
Set xType.RowSource = data9
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set DATA10.Recordset = myRecordSet("select * from PAID_Codes_INV", con)
Set xPaid_code.RowSource = DATA10
xPaid_code.ListField = "Desca"
xPaid_code.BoundColumn = "Code"

Set grid1.DataSource = DATA11
Set grid2.DataSource = DATA12

bEdit = Not retFlag(aSec, "INFORM")
Fixgrd
openCardTable
myUndo
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.Text) Then
     If xCode.Tag = LoadMode Then
        mydefine
    Else
        xCode.Text = ""
    End If
Else
    If (Not (CardTable.EOF)) And xCode.Tag = LoadMode Then
        If CardTable!doc_no = xCode.Text Then
            Exit Sub
        End If
    End If
    
    openCardTable xCode.Text
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
If Not ValidNum(xCode.Text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If Trim(xdesca.Text) = "" Then
    MsgBox "√”„ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If Not xType.MatchedWithList Then
    MsgBox "‰Ê⁄ «·⁄÷ÊÌ… €Ì— „”Ã·", , systemName
    Exit Function
End If
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
SaveText Me, , Array(xCode1.Name, xCode2.Name)
CardTable.Close
CaseTable.Close
Set CardTable = Nothing
Set CaseTable = Nothing
End Sub
Private Sub PrintMembers()
Dim cString As String, temptable As New ADODB.Recordset, loctable As New ADODB.Recordset

contemp.Execute "delete  from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

cString = "SELECT FILE1_50.*, FILE1_51.MEMBER, FILE1_51.DESCA AS DESCA_REL, FILE1_51.DATE_BIRTH AS DATE_BIRTH_REL, FILE1_51.PRINT_DATE AS PRINT_DATE_REL, REL_CODES.DESCA AS REL_CODE_DESCA" & _
          " FROM (FILE1_50 LEFT JOIN FILE1_51 ON FILE1_50.CODE = FILE1_51.MEMBER) LEFT JOIN REL_CODES ON FILE1_51.RELATION = REL_CODES.CODE"

If IsNumeric(xCode1.Text) Then
    cString = cString & turn(cString) & " FILE1_50.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsNumeric(xCode2.Text) Then
    cString = cString & turn(cString) & " FILE1_50.CODE <= " & xCode2.Text
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Do Until loctable.EOF
    temptable.AddNew
    temptable!val1 = loctable!CODE
    temptable!str1 = ArbString(loctable!CODE)
    temptable!str2 = loctable!vip
    temptable!Str3 = loctable!Name
    temptable!str4 = loctable!Title
    temptable!str5 = loctable!Address
    If Not IsNull(loctable!Degree) Then
        temptable!STR6 = GetField("select desca from degree_Codes where code = " & UnCodeSerial(CardTable!Degree, 71))
    End If
    temptable!str7 = loctable!Address
    temptable!str8 = loctable!phone1
    temptable!str9 = loctable!MOBIL
    temptable!Str10 = loctable!Union
    
    temptable!str11 = TurnValue(ArbString(Format(loctable!DATE_BIRTH, "yyyy/mm/dd")))
    temptable!str12 = TurnValue(ArbString(loctable!receipt & ""))
    temptable!str13 = TurnValue(ArbString(Format(loctable!Print_date, "yyyy/mm/dd")))
    
    temptable!val2 = loctable!Member
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
For i = 1 To grid1.rows - 2
    nValue = Val(grid1.TextMatrix(i, 7)) + nValue
Next
If xDied.Value = 0 Then nValue = nValue + nMemValue
xTotal.Caption = Format(nValue, "fixed")
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 FILE1_50.* FROM FILE1_50"
If pCode <> "" Then cWhere = "FILE1_50.CODE " & pSign & addvalue(pCode)

cFilter = ""
If sCode <> "" Then cFilter = "FILE1_50.CODE = " & addvalue(sCode)
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter

If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by FILE1_50.CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by FILE1_50.CODE ASC"
End If

CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If ValidNum(pCode) Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE <= " & pCode & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM FILE1_50 " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If ValidNum(pCode) Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub myUndo()
'On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Text) Then
    openCardTable xCode.Text
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
If Not IsNumeric(xCode.Text) Then Exit Sub
If nPhoto = 0 And xCode.Text Then
    ReplaceFromImage Image, RetPhoto_I(xCode.Text)
Else
    If nPhoto <= grid1.rows - 1 Then
        If IsNumeric(grid1.TextMatrix(nPhoto, 0)) Then
            ReplaceFromImage Image, RetPhoto_I(xCode.Text & "-" & grid1.TextMatrix(nPhoto, 0))
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
        If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                DeletePhoto_I xCode.Text, .TextMatrix(.Row, 0)
                con.BeginTrans
                con.Execute "Delete  from FILE1_51 where id = " & .TextMatrix(.Row, .Cols - 1)
                con.CommitTrans
            End If
            myRemove .Row
            grid1_EnterCell
            On Error Resume Next
            grid1.SetFocus
            Err.Clear
            loadPhoto_Append xCode.Text, grid1.TextMatrix(grid1.Row, 0)
        End If
    ElseIf KeyCode = 13 Then
        CellPos KeyCode, .Row, .Col
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadGrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And (Col <> 1 And Col <> 2) Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
If Not MYVALID Then
    On Error Resume Next
    .SetFocus
    Err.Clear
    myloadGrd
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
'Calctotals
If MyReplace(Row) Then
    If xCode.Tag = DefineMode Then
        myUndo
    End If
    If .TextMatrix(Row, .Cols - 1) = "" Then
        myloadGrd
        .ShowCell .rows - 1, 0
    End If
Else
    myloadGrd
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadGrd
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not ValidNum(.TextMatrix(Row, 0)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 4)) = "" Then Exit Function
If Not IsDate(.TextMatrix(Row, 5)) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
On Error GoTo myerror
If OldRow <> NewRow Then
    loadPhoto_Append xCode.Text, .TextMatrix(NewRow, 0)
End If
End With
Exit Sub
myerror:
xAppendPhoto.Picture = LoadPicture("")
End Sub
Private Sub grid1_EnterCell()
With grid1
If (.Col = 0 And Trim(.TextMatrix(.Row, .Cols - 1)) <> "") Or .Col = 8 Then
    .Editable = flexEDNone
Else
    .Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
grid1_EnterCell
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
    If (Not IsDate(.EditText)) And Trim(.EditText) <> "" Then
        Cancel = True
    Else
        .EditText = Format(.EditText, "yyyy/mm/dd")
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·—ﬁ„|" & "«·ﬁ—«»…|" & "«·‰Ê⁄|" & "«·’›…|" & "«·«”„|" & " «—ÌŒ «·„Ì·«œ|" & " «—ÌŒ «·⁄÷ÊÌ…|" & "„·«ÕŸ« |" & " «—ÌŒ «·ÿ»«⁄…|"
.ColWidth(0) = 800
.ColWidth(1) = 1200
.ColWidth(2) = 1000
.ColWidth(3) = 1400
.ColWidth(4) = 3000
.ColWidth(5) = 1300
.ColWidth(6) = 1300
.ColWidth(7) = 3000
.ColWidth(8) = 1300
.ColWidth(9) = 950
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
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
        aInsert = AddFlag(Empty, "MEMBER", addvalue(xCode.Text))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "RELATION", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "GENDER", addvalue(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "DATE_BIRTH", addDate(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "DATE_BEGIN", addDate(grid1.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(i, 7)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_51")
        Else
            con.Execute addUpdate(aInsert, "FILE1_51", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadGrd()
With grid1
Dim cString As String
cString = "SELECT FILE1_51.CODE,FILE1_51.RELATION,FILE1_51.GENDER,FILE1_51.TITLE,FILE1_51.DESCA,CONVERT(VARCHAR(10),FILE1_51.DATE_BIRTH,111),CONVERT(VARCHAR(10),FILE1_51.DATE_BEGIN,111),FILE1_51.NOTES,CONVERT(VARCHAR(10),FILE1_51.DATE_PRINT,111),FILE1_51.ID " & _
          " FROM FILE1_51"
cString = cString & " WHERE FILE1_51.MEMBER = " & xCode.Text
cString = cString & " ORDER BY FILE1_51.CODE"
Set DATA11.Recordset = myRecordSet(cString, con)
myAddItem
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
For i = 1 To grid1.rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function

Private Sub xCode_Trans_GotFocus()
myGotFocus xCode_Trans
End Sub
Private Sub xCode_Trans_LostFocus()
myLostFocus xCode_Trans
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
Private Sub xMONTHS_GotFocus()
myGotFocus xMonths
End Sub
Private Sub xMONTHS_LostFocus()
myLostFocus xMonths
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
myGotFocus xdesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xdesca
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
Private Sub xjob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xjob_LostFocus()
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
    If KeyCode = 46 And .Row <> .rows - 1 And bEditRecord Then
        If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.BeginTrans
                con.Execute "Delete  from FILE1_55 where id = " & .TextMatrix(.Row, .Cols - 1)
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
'On Error GoTo myerror
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
If Row = .rows - 1 Then
    myAddItem2
End If

If MyReplace(, Row) Then
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
If Not IsDate(.TextMatrix(Row, 0)) Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow2 = True
End Function
Private Sub Grid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid2
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow2(OldRow) Then
        myRemove2 OldRow
    End If
End If
End With
End Sub
Private Sub Grid2_EnterCell()
With grid2
.Editable = flexEDKbdMouse
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
If Col = 0 Then
    If Not IsDate(.EditText) Then
        MsgBox " «—ÌŒ €Ì— „”Ã·"
        Cancel = True
    Else
        .EditText = myFormat_p(.EditText)
    End If
ElseIf Col = 1 Then
    If Trim(.EditText) = "" Then
        MsgBox "«·»Ì«‰ €Ì— „”Ã·"
        Cancel = True
    End If
End If
End With
End Sub
Private Sub fixgrd2()
With grid2
    .FormatString = "«· «—ÌŒ|" & "«·»Ì«‰|" & "«·‰ ÌÃ…|"
    .ColWidth(0) = 1340
    .ColWidth(1) = 7000
    .ColWidth(2) = 6000
    .ColHidden(.Cols - 1) = True
    For i = 0 To .Cols - 2
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub cellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
With grid2
KeyCode = 0
If Col < .Cols - 2 Then
    .Col = Col + 1
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid2, Row + 1, 0, 1)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem2()
With grid2
.AddItem ""
End With
End Sub
Private Function myreplaceGrd2(Row) As Boolean
Dim aInsert As Variant
With grid2
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid2.rows - 2, Row)
        aInsert = AddFlag(Empty, "MEMBER", addvalue(xCode.Text))
        aInsert = AddFlag(aInsert, "DATE", addDate(grid2.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid2.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "PENALITY", addstring(grid2.TextMatrix(i, 2)))
        If grid2.TextMatrix(i, grid2.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_55")
        Else
            con.Execute addUpdate(aInsert, "FILE1_55", "ID = " & grid2.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd2 = True
End Function
Private Sub myloadgrd2()
With grid2
Dim cString As String
cString = "SELECT CONVERT(VARCHAR(10),FILE1_55.DATE,111),FILE1_55.DESCA,FILE1_55.PENALITY,FILE1_55.ID " & _
          " FROM FILE1_55"
cString = cString & " WHERE FILE1_55.MEMBER = " & xCode.Text
cString = cString & " ORDER BY FILE1_55.DATE"
Set DATA12.Recordset = myRecordSet(cString, con)
myAddItem2
fixgrd2
End With
End Sub
Private Sub myRemove2(Row As Long)
grid2.RemoveItem Row
End Sub
Private Function myTrans() As Boolean
Dim aInsert As Variant, aInsert2 As Variant
aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode_Trans.Text))
aInsert = AddFlag(aInsert, "[DESCA]", "[DESCA]")
aInsert = AddFlag(aInsert, "[PHONE]", "[PHONE]")
aInsert = AddFlag(aInsert, "[MOBIL]", "[MOBIL]")
aInsert = AddFlag(aInsert, "[REGION]", "[REGION]")
aInsert = AddFlag(aInsert, "[TITLE]", "[TITLE]")
aInsert = AddFlag(aInsert, "[ADDRESS]", "[ADDRESS]")
aInsert = AddFlag(aInsert, "[JOB]", "[JOB]")
aInsert = AddFlag(aInsert, "[JOB_ADDRESS]", "[JOB_ADDRESS]")
aInsert = AddFlag(aInsert, "[NOTES]", "[NOTES]")
aInsert = AddFlag(aInsert, "[DATE_PRINT]", "[DATE_PRINT]")
aInsert = AddFlag(aInsert, "[DOC_NO]", "[DOC_NO]")
aInsert = AddFlag(aInsert, "[USERNAME]", "[USERNAME]")
aInsert = AddFlag(aInsert, "[TIME]", "[TIME]")
aInsert = AddFlag(aInsert, "[USERNAME2]", "[USERNAME2]")
aInsert = AddFlag(aInsert, "[TIME2]", "[TIME2]")
aInsert = AddFlag(aInsert, "[JOB_CODE]", "[JOB_CODE]")
aInsert = AddFlag(aInsert, "[PHOTO_CODE]", "[PHOTO_CODE]")
aInsert = AddFlag(aInsert, "[DATE_BIRTH]", "[DATE_BIRTH]")
aInsert = AddFlag(aInsert, "[DATE_BEGIN]", "[DATE_BEGIN]")
aInsert = AddFlag(aInsert, "[MEMBERID]", "[MEMBERID]")
aInsert = AddFlag(aInsert, "[ISOLD]", "[ISOLD]")
aInsert = AddFlag(aInsert, "[NOCARD]", "[NOCARD]")
aInsert = AddFlag(aInsert, "[MAIL]", "[MAIL]")
aInsert = AddFlag(aInsert, "[ID_NO]", "[ID_NO]")
aInsert = AddFlag(aInsert, "[GENDER]", "[GENDER]")
aInsert = AddFlag(aInsert, "[SOCIAL]", "[SOCIAL]")
aInsert = AddFlag(aInsert, "[RELIGION]", "[RELIGION]")
aInsert = AddFlag(aInsert, "[DEGREE]", "[DEGREE]")
aInsert = AddFlag(aInsert, "[STATUS]", "[STATUS]")
aInsert = AddFlag(aInsert, "[DATE_TRANS]", "[DATE_TRANS]")
aInsert = AddFlag(aInsert, "[DATE_JOIN]", "[DATE_JOIN]")
aInsert = AddFlag(aInsert, "[DATE_DEGREE]", "[DATE_DEGREE]")
aInsert = AddFlag(aInsert, "[TYPE]", "[TYPE]")
aInsert = AddFlag(aInsert, "[DATE_LAST]", "[DATE_LAST]")
aInsert = AddFlag(aInsert, "[DIED]", "[DIED]")
aInsert = AddFlag(aInsert, "[JOB_DESCA]", "[JOB_DESCA]")
aInsert = AddFlag(aInsert, "[SES_NO]", "[SES_NO]")
aInsert = AddFlag(aInsert, "[COMPANY]", "[COMPANY]")
aInsert = AddFlag(aInsert, "[DROP]", "[DROP]")
aInsert = AddFlag(aInsert, "[FACE]", "[FACE]")
aInsert = AddFlag(aInsert, "[REASON]", "[REASON]")
aInsert = AddFlag(aInsert, "[YEAR_CODE]", "[YEAR_CODE]")
aInsert = AddFlag(aInsert, "[PAID_TYPE]", "[PAID_TYPE]")
aInsert = AddFlag(aInsert, "[REGSTER]", "[REGSTER]")
aInsert = AddFlag(aInsert, "[CODE_INV]", "[CODE]")

aInsert2 = AddFlag(aInsert2, "[MEMBER]", addvalue(xCode_Trans.Text))
aInsert2 = AddFlag(aInsert2, "[CODE]", "[CODE]")
aInsert2 = AddFlag(aInsert2, "[DESCA]", "[DESCA]")
aInsert2 = AddFlag(aInsert2, "[DATE_BIRTH]", "[DATE_BIRTH]")
aInsert2 = AddFlag(aInsert2, "[RELATION]", "[RELATION]")
aInsert2 = AddFlag(aInsert2, "[GENDER]", "[GENDER]")
aInsert2 = AddFlag(aInsert2, "[TITLE]", "[TITLE]")
aInsert2 = AddFlag(aInsert2, "[date_print]", "[date_print]")
aInsert2 = AddFlag(aInsert2, "[NOTES]", "[NOTES]")
aInsert2 = AddFlag(aInsert2, "[MEMBERID]", "[MEMBERID]")
aInsert2 = AddFlag(aInsert2, "[HANDI]", "[HANDI]")
aInsert2 = AddFlag(aInsert2, "[DATE_BEGIN]", "[DATE_BEGIN]")
aInsert2 = AddFlag(aInsert2, "[PHOTO]", "[PHOTO]")
aInsert2 = AddFlag(aInsert2, "[code_Test]", "[code_Test]")
aInsert2 = AddFlag(aInsert2, "[ISMEMBER]", "[ISMEMBER]")

Dim loctable As New ADODB.Recordset
loctable.Open "select * from file1_51 where MEMBER = " & addvalue(xCode.Text), con, adOpenStatic, adLockReadOnly, adCmdText

con.BeginTrans
On Error GoTo myerror
con.Execute addInsertTable(aInsert, "FILE1_10", "FILE1_50", "CODE = " & addvalue(xCode.Text))
con.Execute addInsertTable(aInsert2, "FILE1_11", "FILE1_51", "MEMBER = " & addvalue(xCode.Text))
con.Execute "UPDATE FILE1_50 SET FILE1_50.CODE_TRANS = " & addvalue(xCode_Trans.Text)

Dim fs As New FileSystemObject
If validPhoto(RetPhoto_I(xCode.Text)) Then
    fs.CopyFile RetPhoto_I(xCode.Text), RetPhoto(xCode_Trans.Text)
End If

Do Until loctable.EOF
    If validPhoto(RetAppendPhoto_i(loctable!Member, loctable!CODE)) Then
        fs.CopyFile RetAppendPhoto_i(loctable!Member, loctable!CODE), RetAppendPhoto(xCode_Trans.Text, loctable!CODE)
    End If
    loctable.MoveNext
Loop
con.CommitTrans
myTrans = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function validTrans() As Boolean
If Not ValidNum(xCode_Trans.Text) Then
    MsgBox "—ﬁ„ €Ì— ’«·Õ"
    Exit Function
End If
If Not IsEmpty(GetField("select code from file1_10 where code = " & addvalue(xCode_Trans.Text), con)) Then
    MsgBox "⁄÷Ê ⁄«„· »‰›” «·—ﬁ„"
    Exit Function
End If
validTrans = True
End Function
Private Function loadPhoto(pCode As String) As Boolean
On Error Resume Next
xMemberPhoto.Picture = LoadPicture("")
If Dir(RetPhoto_I(pCode)) <> "" Then xMemberPhoto.Picture = LoadPicture(RetPhoto_I(pCode))
Err.Clear
End Function
Private Function loadPhoto_Append(pCode As String, Optional pAppend As String = "") As Boolean
On Error Resume Next
xAppendPhoto.Picture = LoadPicture("")
If Dir(RetAppendPhoto_i(pCode, pAppend)) <> "" Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto_i(pCode, pAppend))
Err.Clear
End Function


