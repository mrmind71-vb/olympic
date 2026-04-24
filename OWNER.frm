VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form memberfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8790
   ClientLeft      =   615
   ClientTop       =   1320
   ClientWidth     =   15510
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
   Picture         =   "OWNER.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   15510
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Command5"
      Height          =   420
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   8010
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«÷«›… ’Ê— «· Ê«»⁄"
      Height          =   600
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   8055
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "«÷«›… «· Ê«»⁄"
      Height          =   600
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   8010
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«÷«›… «·’Ê—"
      Height          =   600
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   8235
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«÷«›… «·«⁄÷«¡"
      Height          =   600
      Left            =   5580
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   8235
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   7425
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         TabIndex        =   60
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
         Picture         =   "OWNER.frx":0342
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OWNER.frx":2512
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   810
         TabIndex        =   61
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
         Picture         =   "OWNER.frx":465A
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OWNER.frx":6822
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1575
         TabIndex        =   62
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
         Picture         =   "OWNER.frx":8971
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OWNER.frx":AB51
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2340
         TabIndex        =   63
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
         Picture         =   "OWNER.frx":CCAC
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OWNER.frx":EE68
      End
   End
   Begin VB.Frame Frame4 
      Height          =   960
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   4455
      Width           =   5865
      Begin VB.Label xtime2 
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   540
         Width           =   2895
      End
      Begin VB.Label Label22 
         Caption         =   "  ⁄œÌ·"
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
         TabIndex        =   57
         Top             =   630
         Width           =   645
      End
      Begin VB.Label xusername2 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label xtime 
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   "«÷«›…"
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
         TabIndex        =   54
         Top             =   270
         Width           =   600
      End
      Begin VB.Label xusername 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   180
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   3510
      Width           =   5865
      Begin VB.Label Label20 
         Caption         =   "«Œ— ÿ»«⁄…"
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
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   585
         Width           =   825
      End
      Begin VB.Label xDate_print 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label Label15 
         Caption         =   " «Œ— ”œ«œ"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   225
         Width           =   780
      End
      Begin VB.Label xdate_paid 
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
         TabIndex        =   48
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "«Œ— «Ì’«·"
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
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   825
      End
      Begin VB.Label xdoc_no 
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   180
         Width           =   1905
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1680
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3735
      Width           =   9105
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
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1260
         Width           =   7620
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
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   900
         Width           =   7620
      End
      Begin VB.TextBox xJob 
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
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   7620
      End
      Begin VB.TextBox xJob_address 
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
         Left            =   180
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   7620
      End
      Begin VB.Label Label18 
         Caption         =   "—ﬁ„ ﬁÊ„Ì"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   1305
         Width           =   1170
      End
      Begin VB.Label Label16 
         Caption         =   "»—Ìœ «·Ìﬂ —Ê‰Ì"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label10 
         Caption         =   "„ﬂ«‰ «·⁄„·"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "«·ÊŸÌ›…"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1770
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1980
      Width           =   9105
      Begin VB.TextBox xDatebirth 
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "D"
         Top             =   945
         Width           =   2265
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
         Width           =   7665
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
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   7665
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
         Height          =   330
         Left            =   3555
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   945
         Width           =   4245
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   7665
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
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1350
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   315
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label Label13 
         Caption         =   "—ﬁ„ «·«—÷Ì"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.Frame Frame7 
      Height          =   690
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   36
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
         TabIndex        =   18
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
         Picture         =   "OWNER.frx":10FB7
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "OWNER.frx":1331A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "OWNER.frx":15893
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "OWNER.frx":17CFF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Picture         =   "OWNER.frx":1A599
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "OWNER.frx":1CB45
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2850
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   630
      Width           =   5865
      Begin VB.CommandButton cmdScan2 
         BackColor       =   &H00C0FFC0&
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
         Left            =   90
         Picture         =   "OWNER.frx":1F318
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Width           =   1725
      End
      Begin VB.CommandButton cmdScan 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   90
         Picture         =   "OWNER.frx":21A56
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1485
         Width           =   1725
      End
      Begin VB.Image xAppendPhoto 
         Appearance      =   0  'Flat
         Height          =   2610
         Left            =   1890
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1980
      End
      Begin VB.Image xMemberPhoto 
         Appearance      =   0  'Flat
         Height          =   2595
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1800
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
      Height          =   1320
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   675
      Width           =   9105
      Begin VB.CheckBox xNoCard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ ﬂ«—‰ÌÂ"
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
         Height          =   255
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   225
         Width           =   1770
      End
      Begin VB.CheckBox xDied 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "·« Ìÿ»⁄"
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
         Height          =   255
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   180
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.CommandButton cmdJob 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   900
         Width           =   330
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   2220
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   330
         Left            =   5355
         TabIndex        =   3
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
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
         Left            =   3510
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   4290
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
         Left            =   6480
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xUnion_reg 
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   2220
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label7 
         Caption         =   "ﬂÊœ «·⁄÷Ê "
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   30
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   " ·Ì›Ê‰ «·„‰“· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2565
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ «·ﬁÌœ »«·‰ﬁ«»…"
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label12 
         Caption         =   "«·‘⁄»…"
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
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   6750
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   1035
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
      Left            =   7110
      Top             =   7515
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
   Begin ImgXCtrl6.ImgXCtrl imgx1 
      DragIcon        =   "OWNER.frx":24194
      DragMode        =   1  'Automatic
      Height          =   2085
      Left            =   2565
      TabIndex        =   33
      Tag             =   "-1"
      Top             =   1845
      Visible         =   0   'False
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   3678
      BorderStyle     =   1
      AutoZoom        =   -1  'True
      LicenseUserName =   "mrvb71"
      LicenseRegCode  =   "íß“ªß•≤º∂´≠“±®ππ∂´µßZQEH-AOZOOOZT-EFLF6gI"
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   1950
      Left            =   90
      TabIndex        =   15
      Top             =   5445
      Width           =   15270
      _cx             =   26935
      _cy             =   3440
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
      Left            =   3285
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   7515
      Width           =   5505
   End
End
Attribute VB_Name = "memberfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim fs As New FileSystemObject
Public bedit As Boolean, bEditRecord As Boolean
Dim con As New ADODB.Connection
Dim fs As New FileSystemObject
Dim WithEvents twain As ImgXTwain, nphoto As Long
Attribute twain.VB_VarHelpID = -1
Dim cRelStr As String, cGenderStr As String
Dim formMode As Byte
Dim oSearch As New Search3, oSearchRel As New Search3
Dim CardTable As ADODB.Recordset
Public sCode As String
Dim cFilter As String, cFilterLookup As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
bEditRecord = bedit
cmdAdd.Enabled = (nMode = LoadMode And bEditRecord)
CmdDel.Enabled = (nMode = LoadMode And bEditRecord)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = bedit = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = Newflag("file1_10", "code")
xDesca.Text = ""
xTitle.Text = ""
xAddress.Text = ""
xSection.BoundText = ""
xPhone.Text = ""
xMobil.Text = ""
xUnion_reg.Text = ""
xDatebirth.Text = ""
xJob_address.Text = ""
xJob.Text = ""
xId_no.Text = ""
xMail.Text = ""
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")

xDate_print.Caption = ""
xdate_paid.Caption = ""
xDoc_No.Caption = ""
xUserName.Caption = ""
xUserName2.Caption = "'"
xtime.Caption = ""
XTIME2.Caption = ""
xNoCard.Value = 0
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
xRecordNo.Caption = "«÷«›… ”Ã· ÃœÌœ " & "(" & CardTable.RecordCount & ")"
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Unload oSearch
ElseIf ActiveControl.Name = Me.cmdInform_rel.Name Then
    xCode.Text = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 0)
    Unload oSearchRel
End If
myUndo
End Sub
Private Sub MyLoad(Optional bNoGrid As Boolean = False)
Dim aRet As Variant
xCode.Text = CardTable!CODE & ""
xDesca.Text = CardTable!desca & ""
xSection.BoundText = CardTable!Section & ""
xTitle.Text = CardTable!Title & ""
xAddress.Text = CardTable!Address & ""
xDate_print.Caption = Format(CardTable!Date_print, "dd-mm-yyyy")
xDoc_No.Caption = CardTable!doc_no & ""
xPhone.Text = CardTable!Phone & ""
xMobil.Text = CardTable!MOBIL & ""
xUnion_reg.Text = CardTable!Union_reg & ""
xDatebirth.Text = CardTable!DateBirth & ""
xNotes.Text = CardTable!notes & ""
xNoCard.Value = IIf(CardTable!noCard, "1", "0")
xId_no.Text = CardTable!id_no & ""
xMail.Text = CardTable!mail & ""
Handlecontrols LoadMode
xMemberPhoto.Picture = LoadPicture("")
xAppendPhoto.Picture = LoadPicture("")
xUserName.Caption = CardTable!UserName & ""
xUserName2.Caption = CardTable!UserName2 & ""
xtime.Caption = Format(CardTable!Time, "YYYY/MM/DD HH:NN")
XTIME2.Caption = Format(CardTable!Time2, "YYYY/MM/DD HH:NN")
xRecordNo.Caption = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
myloadgrd
grid1.Select 1, 0
'CellPos 13, 0, grid1.Cols - 1
'If grid1.Rows > 2 Then
'    xAppendPhoto.Picture = LoadPicture("")
'    If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(1, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(1, 0)))
'End If
aRet = LastDoc(xCode.Text, con)
xDoc_No.Caption = retFlag(aRet, "doc_no") & ""
xdate_paid.Caption = Format(retFlag(aRet, "date"), "YYYY/M/D")
If bNoGrid Then Exit Sub
On Error Resume Next
grid1.SetFocus
Err.Clear
Exit Sub
'If grid1.Rows > 1 Then
'    If grid1.TextMatrix(1, 0) <> "" Then
'        If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(1, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(1, 0)))
'    End If
'End If
End Sub
Private Function MyReplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[CODE]", addstring(xCode.Text))
aInsert = AddFlag(aInsert, "[DESCA]", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "[TITLE]", addstring(xTitle.Text))
aInsert = AddFlag(aInsert, "[SECTION]", addvalue(xSection.BoundText))
aInsert = AddFlag(aInsert, "[UNION_REG]", addstring(xUnion_reg.Text))
aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(xAddress.Text))
aInsert = AddFlag(aInsert, "[PHONE]", addstring(xPhone.Text))
aInsert = AddFlag(aInsert, "[MOBIL]", addstring(xMobil.Text))
aInsert = AddFlag(aInsert, "[DATEBIRTH]", addDate(xDatebirth.Text))
aInsert = AddFlag(aInsert, "[JOB]", addstring(xJob.Text))
aInsert = AddFlag(aInsert, "[JOB_ADDRESS]", addstring(xJob_address.Text))
aInsert = AddFlag(aInsert, "[NOCARD]", xNoCard.Value)
aInsert = AddFlag(aInsert, "[MAIL]", addstring(xMail.Text))
aInsert = AddFlag(aInsert, "[ID_NO]", addstring(xId_no.Text))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xCode.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
con.BeginTrans
On Error GoTo myerror
If xCode.Tag = DefineMode Then
    con.Execute addInsert(aInsert, "FILE1_10")
Else
    con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & xCode.Text)
End If
myreplaceGrd Row
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
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If grid1.Rows > 2 Then
        MsgBox "«·⁄÷Ê ·Â  Ê«»⁄ ÌÃ» Õ–›Â„ «Ê·«"
        Exit Sub
    End If
    con.BeginTrans
    con.Execute "Delete  From FILE1_10 Where code = " & xCode.Text
    Dim fs As New FileSystemObject
    If fs.FileExists(RetPhoto(xCode.Text)) Then
        fs.DeleteFile RetPhoto(xCode.Text)
    End If
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & xCode.Text, , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
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
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
MemberLookupAll Me, oSearch, cFilter
End Sub
Private Sub cmdInform_rel_Click()
relLookupAll Me, oSearchRel
End Sub

Private Sub cmdJob_Click()
Dim oFlagfrm As New flag_mainfrm, sBoundText As String
sBoundText = xSection.BoundText
oFlagfrm.sTable = "SECTION_CODES"
oFlagfrm.sCaption = "«·‘⁄»…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
Set DATA1.Recordset = myRecordSet("select * from section_Codes", con)
xSection.BoundText = sBoundText
If Not xSection.MatchedWithList Then xSection.BoundText = ""
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

Private Sub CmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub cmdScan_Click()
Scan.sCode = xCode.Text
Scan.Show 1
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
End If
MyLoad
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub cmdScan2_Click()
nphoto = 0
ScanImage
On Error Resume Next
If validPhoto(RetPhoto(xCode.Text)) Then xMemberPhoto.Picture = LoadPicture(RetPhoto(xCode.Text))
If grid1.TextMatrix(grid1.Row, 0) <> "" And grid1.Row <> 0 Then
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.Row)) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
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
'loctable.Open "select * from file1_10 where NEWDATA = true", con, adOpenStatic, adLockReadOnly, adCmdText
loctable.Open "select * from file1_10", con, adOpenStatic, adLockReadOnly, adCmdText
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
            fs.CopyFile cFile, RetPhoto(loctable!CODE)
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
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_10", conmdb, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!desca & ""))
    aInsert = AddFlag(aInsert, "SECTION", addvalue(loctable!Section & ""))
    aInsert = AddFlag(aInsert, "dateBirth", addDate(Format(loctable!DateBirth, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "union_reg", addstring(loctable!Union_reg & ""))
    aInsert = AddFlag(aInsert, "NOTES", addstring(loctable!notes & ""))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address & ""))
    aInsert = AddFlag(aInsert, "JOB_CODE", addstring(loctable!JOB_CODE & ""))
    aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!Phone & ""))
    aInsert = AddFlag(aInsert, "MOBIL", addstring(loctable!MOBIL & ""))
    aInsert = AddFlag(aInsert, "PHOTO_CODE", addstring(loctable!PHOTO_CODE & ""))
    con.Execute addInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conmdb.Close
Set conmdb = Nothing
Exit Sub
myerror:
MsgBox Err.Description
End Sub
Private Sub addRelation()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM FILE1_11", conmdb, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!desca & ""))
    aInsert = AddFlag(aInsert, "dateBirth", addDate(Format(loctable!DateBirth, "YYYY-MM-DD")))
    aInsert = AddFlag(aInsert, "RELATION", addvalue(loctable!Relation & ""))
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
conmdb.Close
Set conmdb = Nothing
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
loctable.Open "select * from file1_11", con, adOpenStatic, adLockReadOnly, adCmdText
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
            fs.CopyFile cFile, RetAppendPhoto(loctable!Member, loctable!CODE)
        End If
    End If
    loctable.MoveNext
Loop
MsgBox "Done"
End Sub
Private Sub Command5_Click()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & App.Path & "\MDB\DATA.mdb"
Dim cFile As String

loctable.Open "SELECT * FROM RELATION_CODES", conmdb, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!desca & ""))
    con.Execute addInsert(aInsert, "RELATION_CODES")
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
conmdb.Close
Set conmdb = Nothing
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Command6_Click()
Dim fs As New FileSystemObject, f, f1, fc, s
Set f = fs.GetFolder(App.Path & "\photo_fix")
Set fc = f.Files
nCount = fc.Count
Dim cString As String, i As Long
bCrypt = True
For Each f1 In fc
    i = i + 1
    Me.Caption = i
    If InStr(1, LCase(App.Path & "\photo_fix\" & f1.Name), "jpg") <> 0 Then
        cFile = RetPhoto(Replace(LCase(f1.Name), ".jpg", ""))
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
    bNoCard = Not validPhoto(RetPhoto(loctable!CODE))
    aInsert = AddFlag(Empty, "nocard", IIf(bNoCard, "1", "0"))
    con.Execute addUpdate(aInsert, "FILE1_10", "CODE = " & loctable!CODE)
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
Exit Sub
myerror:
MsgBox Err.Description
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
'makeMyLoad Me
'makeMyReplace Me
'makeMyDefine Me
'MFocus Me
'LostFocus Me
'makeMyValidate MeLoadText Me
'LoadText Me
'nMemValue = Val(GetDesca("Select value from rel_codes where code = -1"))
openCon con

cRelStr = StrList("Select Code,Desca From relation_codes where code <> 0")
cGenderStr = StrList("Select Code,Desca From gender_codes order by Code")

Set DATA1.Recordset = myRecordSet("select * from section_Codes", con)
Set xSection.RowSource = DATA1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set grid1.DataSource = DATA11
'data10.ConnectionString = con.ConnectionString

bedit = True
Fixgrd
openCardTable
myUndo
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not ValidInt(xCode.Text) Then Exit Sub
CardTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    MyLoad True
ElseIf xCode.Tag = LoadMode Then
    mydefine
End If
End Sub
Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not ValidInt(xCode.Text) Then
    If Not igMsg Then MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

If xDesca.Text = "" Then
    If Not igMsg Then MsgBox "√”„ «·⁄÷Ê €Ì— „”Ã·", , systemName
    Exit Function
End If

For i = 1 To grid1.Rows - 2
    If Not ValidInt(grid1.TextMatrix(i, 0)) Then
        MsgBox "ﬂÊœ «· «»⁄ €Ì— „”Ã·"
        Exit Function
    End If

   
    If Not ValidInt(grid1.TextMatrix(i, 1)) Then
        MsgBox "‰Ê⁄ «· »⁄Ì… €Ì— „”Ã·…"
        Exit Function
    End If
Next
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
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
'If Not bSupermode Then Exit Sub
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            Dim fs As New FileSystemObject
            If fs.FileExists(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0))) Then
                fs.DeleteFile RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0))
            End If
            con.BeginTrans
            con.Execute "Delete  from file1_11 where id = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
            xAppendPhoto.Picture = LoadPicture("")
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub PrintMembers()
Dim cString As String, temptable As New ADODB.Recordset, loctable As New ADODB.Recordset

contemp.Execute "delete  from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

cString = "SELECT FILE1_10.*, FILE1_11.MEMBER, FILE1_11.DESCA AS DESCA_REL, FILE1_11.DATEBIRTH AS DATEBIRTH_REL, FILE1_11.PRINT_DATE AS PRINT_DATE_REL, REL_CODES.DESCA AS REL_CODE_DESCA" & _
          " FROM (FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER) LEFT JOIN REL_CODES ON FILE1_11.RELATION = REL_CODES.CODE"

If IsNumeric(xCode1.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsNumeric(xCode2.Text) Then
    cString = cString & turn(cString) & " File1_10.CODE <= " & xCode2.Text
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
        temptable!str6 = GetField("select desca from degree_Codes where code = " & UnCodeSerial(CardTable!Degree, 71))
    End If
    temptable!str7 = loctable!Address
    temptable!str8 = loctable!phone1
    temptable!str9 = loctable!MOBIL
    temptable!Str10 = loctable!Union
    
    temptable!str11 = TurnValue(ArbString(Format(loctable!DateBirth, "yyyy/mm/dd")))
    temptable!str12 = TurnValue(ArbString(loctable!receipt & ""))
    temptable!str13 = TurnValue(ArbString(Format(loctable!Print_date, "yyyy/mm/dd")))
    
    temptable!val2 = loctable!Member
    temptable!str16 = loctable!Desca_rel
    temptable!str17 = loctable!REL_CODE_DESCA
    temptable!str18 = TurnValue(ArbString(Format(loctable!Print_date_rel, "yyyy/mm/dd")))
    temptable!str19 = TurnValue(ArbString(Format(loctable!DateBirth_rel, "yyyy/mm/dd")))
    temptable!Val3 = retPaid(loctable!CODE)
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
Private Sub Calctotals()
Dim nValue
For i = 1 To grid1.Rows - 2
    nValue = Val(grid1.TextMatrix(i, 7)) + nValue
Next
If xDied.Value = 0 Then nValue = nValue + nMemValue
xTotal.Caption = Format(nValue, "fixed")
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE1_10.* FROM FILE1_10"
If IsNumeric(sCode) Then cString = cString & turn(cString) & " CODE = " & sCode
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY  FILE1_10.code"
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
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hWnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
twain.Acquire False, Me.hWnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
If Not IsNumeric(xCode.Text) Then Exit Sub
If nphoto = 0 And xCode.Text Then
    ReplaceFromImage Image, RetPhoto(xCode.Text)
Else
    If nphoto <= grid1.Rows - 1 Then
        If IsNumeric(grid1.TextMatrix(nphoto, 0)) Then
            ReplaceFromImage Image, RetPhoto(xCode.Text & "-" & grid1.TextMatrix(nphoto, 0))
        End If
    End If
End If
nphoto = nphoto + 1
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
Private Function retPaid(pMember As String) As Double
Dim aRet As Variant, cString As String
aRet = GetField("Select code from file1_10 where (not died) and  code = " & pMember)
If Not IsEmpty(aRet) Then retPaid = nMemValue

cString = "SELECT SUM(REL_CODES.[VALUE])" & _
          " FROM FILE1_11 INNER JOIN REL_CODES ON FILE1_11.RELATION = REL_CODES.CODE WHERE FILE1_11.MEMBER= " & pMember
aRet = GetField(cString)
If Not IsEmpty(cString) Then
    retPaid = retPaid + Val(aRet & "")
End If
End Function

Private Sub xjob_GotFocus()
myGotFocus xJob
End Sub
Private Sub xjob_LostFocus()
myLostFocus xJob
End Sub
Private Sub xJob_address_GotFocus()
myGotFocus xJob_address
End Sub
Private Sub xJob_address_LostFocus()
myLostFocus xJob_address
End Sub
Private Sub xDatebirth_GotFocus()
myGotFocus xDatebirth
End Sub
Private Sub xDatebirth_LostFocus()
myLostFocus xDatebirth
myValidDate xDatebirth
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xNotes_LostFocus()
myLostFocus xNotes
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
Private Sub xAddress_GotFocus()
myGotFocus xAddress
End Sub
Private Sub xAddress_LostFocus()
myLostFocus xAddress
End Sub
Private Sub xTitle_GotFocus()
myGotFocus xTitle
End Sub
Private Sub xTitle_LostFocus()
myLostFocus xTitle
End Sub
Private Sub xSection_GotFocus()
myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
myLostFocus xSection
If Not xSection.MatchedWithList Then xSection.BoundText = ""
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

Private Sub xUnion_reg_GotFocus()
myGotFocus xUnion_reg
End Sub
Private Sub xUnion_reg_LostFocus()
myLostFocus xUnion_reg
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And (Col <> 1 And Col <> 2) Then CellPos KeyCode, Row, Col
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
With grid1
'If Row = .Rows - 1 Then
'    xAppendPhoto.Picture = LoadPicture("")
'    If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(grid1.Row, 0)))
'End If
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    myAddItem
End If
'Calctotals
If MyReplace(Row) Then
    Handlecontrols LoadMode
    If grid1.TextMatrix(Row, .Cols - 1) = "" Then
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
If Not MYVALID(True) Then Exit Function
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
'If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function
If Trim(.TextMatrix(Row, 4)) = "" Then Exit Function

'If Trim(.TextMatrix(Row, 2)) = "" Then Exit Function

End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow < 1 Then Exit Sub
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
    End If
End If
On Error GoTo myerror
If OldRow <> NewRow Then
    xAppendPhoto.Picture = LoadPicture("")
    If validPhoto(RetAppendPhoto(xCode.Text, grid1.TextMatrix(NewRow, 0))) Then xAppendPhoto.Picture = LoadPicture(RetAppendPhoto(xCode.Text, grid1.TextMatrix(NewRow, 0)))
End If
Exit Sub
myerror:
xAppendPhoto.Picture = LoadPicture("")
End Sub
Private Sub Grid1_EnterCell()
With grid1
If (.Col = 0 And Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "") Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End With
End Sub
Private Sub Grid1_GotFocus()
'CellPos 13, grid1.Rows - 2, grid1.Cols - 1
Grid1_EnterCell
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If OldRow < 1 Then Exit Sub
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    myRemove grid1.Row
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then
        If grid1.Row = grid1.Rows - 1 Then Exit Sub
        MsgBox "ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    Else
        nFound = FoundOtheritem(grid1, Row, 0, Trim(grid1.EditText))
        If nFound <> -1 Then
            MsgBox "«·ﬂÊœ „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
            Cancel = True
            Exit Sub
        End If
    End If
ElseIf Col = 2 Then
    If Trim(grid1.EditText) = "" Then
        Cancel = True
    End If
ElseIf Col = 5 Then
    If Not IsDate(grid1.EditText) Then
        Cancel = True
    Else
        grid1.EditText = Format(grid1.EditText, "yyyy/mm/dd")
    End If
End If
End With
End Sub
Private Sub Fixgrd()
With grid1
.FormatString = "«·—ﬁ„|" & "«·ﬁ—«»…|" & "«·‰Ê⁄|" & "«·’›…|" & "«·«”„|" & " «—ÌŒ «·„Ì·«œ|" & "„·«ÕŸ« |" & "»œÊ‰ ﬂ«—‰ÌÂ|" & "–ÊÌ «Õ Ì«Ã« |"
.ColWidth(0) = 800
.ColWidth(1) = 1800
.ColWidth(2) = 900
.ColWidth(3) = 1200
.ColWidth(4) = 4000
.ColWidth(5) = 1500
.ColWidth(6) = 2500
.ColWidth(7) = 1000
.ColWidth(8) = 1000
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
    .Col = Col + 1 + IIf(Col = 0, 1, 0)
ElseIf Row < .Rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 3)
    .ShowCell Row + 1, 0
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
Dim aInsert As Variant
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "MEMBER", addvalue(xCode.Text))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "RELATION", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "GENDER", addvalue(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "DATEBIRTH", addDate(grid1.TextMatrix(i, 5)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(grid1.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "NOCARD", IIf(Val(grid1.TextMatrix(i, 7)) = 0, "0", "1"))
        aInsert = AddFlag(aInsert, "HANDI", IIf(Val(grid1.TextMatrix(i, 8)) = 0, "0", "1"))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE1_11")
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
Dim cString As String
cString = "SELECT FILE1_11.CODE,FILE1_11.RELATION,FILE1_11.GENDER,FILE1_11.TITLE,FILE1_11.DESCA,CONVERT(VARCHAR(10),FILE1_11.DATEBIRTH,111),FILE1_11.NOTES,FILE1_11.NOCARD,FILE1_11.HANDI,FILE1_11.ID " & _
          " FROM FILE1_11"
cString = cString & turn(cString) & "FILE1_11.MEMBER = " & xCode.Text
cString = cString & " ORDER BY CODE"
Set DATA11.Recordset = myRecordSet(cString, con)
myAddItem
Fixgrd
End With
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
'Calctotals
End Sub
Private Function FoundOtheritem(grid1 As Variant, nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function


Private Sub xMail_GotFocus()
myGotFocus xMail
End Sub
Private Sub xMail_LostFocus()
myLostFocus xMail
End Sub
Private Sub xID_NO_GotFocus()
myGotFocus xId_no
End Sub
Private Sub xID_NO_LostFocus()
myLostFocus xId_no
End Sub

