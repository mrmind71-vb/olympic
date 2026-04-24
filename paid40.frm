VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paidfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ì’«·«  ”œ«œ"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   20250
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   20250
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame20 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10215
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   -45
      Width           =   3750
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "«€·«ﬁ „” ‰œ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   225
         Width           =   1365
      End
      Begin Threed.SSCommand cmdClosePeriod 
         Height          =   420
         Left            =   90
         TabIndex        =   68
         Top             =   225
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   741
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
         Caption         =   "«€·«ﬁ › —…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   0
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1125
      Width           =   1950
      Begin Threed.SSCommand cmdAddItems 
         Height          =   1050
         Left            =   45
         TabIndex        =   46
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   1852
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
         Caption         =   "√÷«›… »‰Êœ «·„ÿ«·»…"
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.CheckBox xCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ".«·„Ê”„ «·Õ«·Ì ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8730
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7650
      Value           =   1  'Checked
      Width           =   2445
   End
   Begin VB.CheckBox xAdded 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7875
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8190
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   13995
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   6180
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   4995
         TabIndex        =   30
         TabStop         =   0   'False
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
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "paid40.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":23CB
      End
      Begin Threed.SSCommand cmdNewInv 
         Height          =   510
         Left            =   3735
         TabIndex        =   31
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
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "paid40.frx":4474
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":647C
      End
      Begin Threed.SSCommand cmddel 
         Height          =   510
         Left            =   2475
         TabIndex        =   32
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
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "paid40.frx":8433
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":ABCF
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   33
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
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "paid40.frx":D063
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   510
         Left            =   1260
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
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
         Picture         =   "paid40.frx":F386
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid40.frx":116FC
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   12795
      Begin VB.TextBox xYears 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4050
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Tag             =   "N"
         Top             =   1035
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox xForm_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5895
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   2490
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10125
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "N"
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9765
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9765
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   8640
         TabIndex        =   17
         Top             =   900
         Width           =   2760
         _ExtentX        =   4868
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
      Begin Threed.SSCommand cmdData 
         Height          =   375
         Left            =   5265
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
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
         Caption         =   "»Ì«‰«  «·⁄÷Ê"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":1387F
      End
      Begin Threed.SSCommand cmdYearChange 
         Height          =   375
         Left            =   5265
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   495
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   661
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
         Caption         =   " ⁄œÌ·"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ…"
         Height          =   285
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   585
         Width           =   1485
      End
      Begin VB.Label xDate_Begin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   540
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label xdoc_no_zero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   675
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label xType_Member 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   315
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ ”‰Ê«  ·„  ”œœ"
         Height          =   240
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1305
         Width           =   1665
      End
      Begin VB.Label xUnPaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1260
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label xType_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   900
         Width           =   3345
      End
      Begin VB.Label xLast_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   20
         Top             =   900
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«Œ— ”œ«œ"
         Height          =   240
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰Ê⁄ «·„ÿ«·»…"
         Height          =   285
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·ﬁ”Ì„…"
         Height          =   240
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   930
      End
      Begin VB.Label xYear_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   2490
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê”„"
         Height          =   240
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   765
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1305
         Width           =   1125
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1260
         Width           =   3525
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
         Height          =   270
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   510
      End
   End
   Begin VB.Frame FRAME_CUR 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   4
      Left            =   5580
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1125
      Width           =   1770
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   45
         TabIndex        =   34
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   900
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
         Picture         =   "paid40.frx":15B6C
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":18491
      End
      Begin Threed.SSCommand cmdUndo 
         Height          =   510
         Left            =   45
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   675
         Width           =   1680
         _ExtentX        =   2963
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
         Picture         =   "paid40.frx":1ACE5
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "paid40.frx":1CE45
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   1980
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      CopiesToPrinter =   2
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   2070
      Top             =   8685
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5805
      Top             =   8190
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   50
      Top             =   9090
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
      Height          =   5145
      Left            =   3600
      TabIndex        =   57
      Top             =   2385
      Width           =   16575
      _cx             =   29236
      _cy             =   9075
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
      ForeColorSel    =   -2147483630
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
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
      WordWrap        =   -1  'True
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   7350
      Left            =   45
      TabIndex        =   58
      Top             =   135
      Width           =   3480
      _cx             =   6138
      _cy             =   12965
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
      ForeColorSel    =   -2147483630
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Frame FRAME10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   11250
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   7515
      Width           =   8880
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬁÌ„… „÷«›…"
         Height          =   240
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label xTotal_Tax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label xTotal_late 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  „ √Œ—…"
         Height          =   240
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label xTotal_Year 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«‘ —«ﬂ«  «·”‰…"
         Height          =   240
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   1275
      End
      Begin VB.Label xTotal_year_other 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·≈Ã„«·Ï"
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "€—«„…  √ŒÌ—"
         Height          =   240
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   7425
      Width           =   3480
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2610
         TabIndex        =   37
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
         Picture         =   "paid40.frx":1F132
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid40.frx":212D9
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1710
         TabIndex        =   38
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
         Picture         =   "paid40.frx":23320
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid40.frx":2540B
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   855
         TabIndex        =   39
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
         Picture         =   "paid40.frx":27405
         Caption         =   "·«Õﬁ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid40.frx":29516
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   780
         _ExtentX        =   1376
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
         Picture         =   "paid40.frx":2B510
         Caption         =   "√ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "paid40.frx":2D734
      End
   End
   Begin VB.Label xYears_desca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7155
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   7650
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label xYear_code 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3825
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   7740
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label xBranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8280
      Visible         =   0   'False
      Width           =   2490
   End
End
Attribute VB_Name = "paidfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String, sCode As String, bNew As Boolean, sType As String
Dim cList As String, cFilter As String, bClosed As Boolean
Dim CardTable As ADODB.Recordset, loctable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim oSearchDoc As New Search3, oSearchMember As New Search, oSearchItems As New Search3, oSearchRel As New Search3, oSearchYearChange As New Search_empty
Dim bEditRecord As Boolean, bAct As Boolean, aPen As Variant
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, cdef_Box As String
Dim formMode
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function MyReplace(Optional Row As Long = -1, Optional bNewOnly As Boolean = False) As Boolean
Dim aInsert As Variant, I As Integer
aInsert = AddFlag(Empty, "[DATE]", addDate(xdate.Text))
aInsert = AddFlag(aInsert, "[CODE]", addvalue(xCode.Text))
aInsert = AddFlag(aInsert, "[TYPE]", addvalue(xType.BoundText))
aInsert = AddFlag(aInsert, "[YEAR_CODE]", addstring(xYear_code.Caption))
aInsert = AddFlag(aInsert, "[YEARS]", Val(xYears.Text))
aInsert = AddFlag(aInsert, "FORM_NO", addstring(xForm_No.Text))
aInsert = AddFlag(aInsert, IIf(xDoc_no.Tag = DefineMode, "[USERNAME]", "[USERNAME2]"), addstring(cUserName))
aInsert = AddFlag(aInsert, IIf(xDoc_no.Tag = DefineMode, "[TIME]", "[TIME2]"), "getdate()")
aInsert = AddFlag(aInsert, IIf(xDoc_no.Tag = DefineMode, "[USERCODE]", "[USERCODE2]"), addvalue(nUsercode))
con.BeginTrans
On Error GoTo myerror
If xDoc_no.Tag = DefineMode Then
    xDoc_no.Text = Newflag("FILE6_20H", "DOC_NO")
    aInsert = AddFlag(Date, "[DATE_ISSUE]", addDate(xdate.Text))
    aInsert = AddFlag(aInsert, "[YEARS_DESCA]", addstring(xYear_code.Caption))
    aInsert = AddFlag(aInsert, "DOC_NO", addvalue(xDoc_no.Text))
    con.Execute addInsert(aInsert, "FILE6_20H")
Else
    con.Execute addUpdate(aInsert, "FILE6_20H", "doc_no = " & addstring(xDoc_no.Text))
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
Private Sub myreplaceGrd(Row As Long)
Dim aInsert As Variant
'.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "≈Ã„«·Ì|" & "‰”»… Œ’„|" & "ﬁÌ„… Œ’„|" & "‰”»… ﬁÌ„… „÷«›…|" & "ﬁÌ„…  „÷«›…|" & "‰”»… €—«„…|" & "ﬁÌ„… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"

With grid1
    For I = IIf(Row = -1, 1, Row) To IIf(Row = -1, .rows - 2, Row)
        aInsert = AddFlag(Empty, "DOC_NO", addstring(xDoc_no.Text))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "VALUE", Val(.TextMatrix(I, 2)))
        aInsert = AddFlag(aInsert, "QUANT", Val(.TextMatrix(I, 3)))
        aInsert = AddFlag(aInsert, "DISCOUNT_RATE", Val(.TextMatrix(I, 4)))
        aInsert = AddFlag(aInsert, "TAX_RATE", Val(.TextMatrix(I, 5)))
        aInsert = AddFlag(aInsert, "LATE_RATE", Val(.TextMatrix(I, 6)))
        aInsert = AddFlag(aInsert, "NOTES", addstring(.TextMatrix(I, 8)))
        If .TextMatrix(I, .Cols - 1) = "" Then
            If grid2.Row < 0 Then
                aInsert = AddFlag(aInsert, "YEAR_CODE", addvalue(xYear_code.Caption))
            ElseIf Not ValidNum(grid2.TextMatrix(grid2.Row, grid2.Cols - 1)) Then
                aInsert = AddFlag(aInsert, "YEAR_CODE", addvalue(xYear_code.Caption))
            Else
                aInsert = AddFlag(aInsert, "YEAR_CODE", addvalue(grid2.TextMatrix(grid2.Row, grid2.Cols - 1)))
            End If
            con.Execute addInsert(aInsert, "FILE6_20")
        Else
            con.Execute addUpdate(aInsert, "FILE6_20", "ID = " & .TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.Text = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    xCodeDesca.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 1)
    Unload oSearchMember
ElseIf ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 1) = oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 1)
        GrdDesc oSearchItems.grid1.TextMatrix(oSearchItems.grid1.Row, 0), grid1.Row
        grid1_AfterEdit grid1.Row, grid1.Col
        Unload oSearchItems
        CellPos 13, grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = Me.cmdYearChange.Name Then
    If ChangeYear(oSearchYearChange.grid1.TextMatrix(oSearchYearChange.grid1.Row, 0)) Then
        Inform " „  €ÌÌ— »‰Ã«Õ"
        myUndo
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    xDoc_no.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
End If
End Sub
Private Sub cmd_closed_Click()
con.BeginTrans
On Error GoTo myerror
con.Execute " update " & cFileHeader & " set CLOSED = " & IIf(xClosed.Value = 1, "0", "1") & " WHERE doc_no = " & MyParn(xDoc_no.Text)
con.CommitTrans
Err.Clear
'openCardTable
myUndo
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdAddItems_Click()
myAdditems
End Sub
Private Function myAdditems() As Boolean
Dim nYears As Long, nFirstYear As Integer, aRet As Variant
If Not ValidNum(xCode.Text) Then
    MsgBox "ﬂÊœ «·⁄÷Ê €Ì— ’ÕÌÕ"
    Exit Function
End If

If Not IsDate(xdate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If

aRet = DocSameDay(xCode.Text, xType.BoundText, xdate.Text, con)
If Not IsEmpty(aRet) Then
    MsgBox "„” ‰œ »‰›” ‰Ê⁄ «·„ÿ«·»… »‰›” «·ÌÊ„ —ﬁ„ " & aRet
    xDoc_no.Text = aRet
    xDoc_No_LostFocus
    Exit Function
End If

aRet = addPayment(xCode.Text, myFormat(xdate.Text), xType.BoundText, con)
If IsEmpty(retFlag(aRet, "error")) Then
    If Not IsEmpty(retFlag(aRet, "msg")) Then
        MsgBox retFlag(aRet, "msg")
    End If
    cInsert = retFlag(aRet, "sql") & ""
    If cInsert <> "" Then
        con.BeginTrans
        On Error GoTo myerror
        con.Execute cInsert
        con.CommitTrans
        xDoc_no.Text = retFlag(aRet, "doc_no")
        myUndo
    End If
Else
    MsgBox retFlag(aRet, "error")
End If

'Dim I As Integer
'For I = 0 To grid1.UBound
'    grid1(I).rows = 1
'    If I > 1 Then
'        SSTab1.TabCaption(I) = ""
'        SSTab1.TabVisible(I) = False
'    End If
'    myAddItem I
'Next
'
'If findRows(aPaidTypes, "code", xType.BoundText, "is_paid", , False) Then
'    xYear_Desca.Caption = retFlag(aYear, "code")
'    nFirstYear = retFlag(aYear, "CODE")
'    For I = 0 To nYears - 1
'        If I = 0 Then
'            xYear_code.Caption = nFirstYear
'        ElseIf I = 1 Then
'            xyear_code1.Caption = nFirstYear - 1
'        ElseIf I = 2 Then
'            xYear_code2.Caption = nFirstYear - 2
'        ElseIf I = 3 Then
'            xYear_code3.Caption = nFirstYear - 3
'        End If
'        SSTab1.TabCaption(I) = Year_Load(nFirstYear - I, "desca")
'        SSTab1.TabVisible(I) = True
'        addPaidItems I, nFirstYear - I
'    Next
'Else
'    xYear_code.Caption = nFirstYear
'    xYear_Desca.Caption = retFlag(aYear, "code")
'    SSTab1.TabCaption(0) = Year_Load(nFirstYear, "desca")
'    SSTab1.TabVisible(0) = True
'    addPaidItems 0, nFirstYear
'End If
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function AddMemberData(aMember As Variant, Index As Variant) As Boolean
Dim nAge As Integer, nGender As Integer
If IsDate(retFlag(aMember, "DATE_BIRTH") & "") Then
   nAge = Age(myFormat(retFlag(aMember, "DATE_BIRTH")), myFormat(xdate.Text)) - Index
Else
   nAge = 1
End If

If Val(loctable!Age1 & "") > nAge And Val(loctable!Age1) <> 0 Then Exit Function
If Val(loctable!Age2 & "") < nAge And Val(loctable!Age2 & "") <> 0 Then Exit Function
If (Not IsNull(loctable!GENDER)) Then
    nGender = TurnValue(retFlag(aMember, "Gender", 1), Null, 1)
    If nGender <> loctable!GENDER Then Exit Function
End If
AddMemberData = True
End Function
Private Function addRelation(nRelation As Integer, ByVal aMeet As Variant) As Integer
Dim myRecordSet As New ADODB.Recordset
Dim nAge As Integer, nGender As Integer
cString = " SELECT [CODE],[DATE_BIRTH],COALESCE(GENDER,1) From FILE1_11"
cString = cString & " where relation = " & nRelation
cString = cString & " AND MEMBER = " & xCode.Text
If Not IsNull(loctable!GENDER) Then cString = cString & " AND COALESCE(GENDER,1) = " & loctable!GENDER
myRecordSet.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until myRecordSet.EOF
    If IsDate(myRecordSet!DATE_BIRTH & "") Then
       nAge = Age(myFormat(myRecordSet!DATE_BIRTH), myFormat(xdate.Text)) - Index
    Else
       nAge = 99
    End If
    If (nAge1 >= Val(loctable!Age1 & "") Or Val(loctable!Age1 & "") = 0) And (nAge2 <= Val(loctable!Age2 & "") Or Val(loctable!Age2 & "") = 0) Then
        addRelation = addRelation + 1
        If nRelation = 1 Then
            aMeet = AddFlag(Empty, loctable!code)
            aMeet = AddFlag(aMeet, loctable!code)
        End If
    End If
    myRecordSet.MoveNext
Loop
myRecordSet.Close
Set myRecordSet = Nothing
End Function

Private Sub cmdClosePeriod_Click()
closefrm.sFile = "FILE6_20H"
closefrm.Show 1
myUndo
End Sub

Private Sub cmdData_Click()
Dim oMember As New memberfrm
If ValidNum(xCode.Text) Then
    oMember.sCode = xCode.Text
    oMember.Show
End If
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE6_20 where Doc_No = " & xDoc_no.Text
    con.Execute "Delete From FILE6_20H where Doc_No = " & xDoc_no.Text
    con.CommitTrans
    If sDoc_no <> "" Then
        Unload Me
        Exit Sub
    End If
    
    'openCardTable xdoc_no_zero.Caption, "<="
    openCardTable xDoc_no.Text, "<="
    If CardTable.EOF Then openCardTable , ">"
    If CardTable.EOF Then
        mydefine
    Else
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Function retAll(aMember As Variant) As Integer
retAll = IIf(retFlag(aMember, "Died", False), 0, 1)
Dim cString As String
cString = "SELECT SUM(1) FROM FILE1_11"
cString = cString & " WHERE FILE1_11.MEMBER = " & xCode.Text
retAll = retAll + Val(GetField(cString, con) & "")
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(4, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_20H.DOC_NO, FILE6_20H.FORM_NO,PAID_TYPES.DESCA,CONVERT(VARCHAR(10),FILE6_20H.DATE,111),FILE6_20H.YEAR_CODE, FILE1_10.DESCA" & _
          "  FROM  FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE"
cString = cString & " WHERE FILE6_20H.OLD = 0"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_20H.DATE,FILE6_20H.YEAR_CODE,FILE6_20H.Doc_No"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·«” „«—…- «—ÌŒ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(0, 1) = "(%%FILE1_10.Desca%% or **FILE6_20H.FORM_NO**" & _
                  " OR ##FILE6_20.Date##)"

listarray(1, 0) = "ﬂÊœ «·⁄÷Ê"
listarray(1, 1) = "(**FILE6_20H.CODE**)"

listarray(2, 0) = "—ﬁ„ «·„” ‰œ"
listarray(2, 1) = "(**FILE6_20H.DOC_NO**)"

listarray(3, 0) = "”‰… «·”œ«œ"
listarray(3, 1) = "(**FILE6_20H.YEAR_CODE**)"

listarray(4, 0) = "‰Ê⁄ «·„ÿ«·»…"
listarray(4, 1) = "(**FILE6_20H.[TYPE]**)"
listarray(4, 2) = "SELECT CODE,DESCA FROM PAID_TYPES"
listarray(4, 3) = "CODE"
listarray(4, 4) = "DESCA"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·«Ì’«·"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "‰Ê⁄ «·„” ‰œ"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(3, 1) = 1350

GrdArray(4, 0) = "”‰… «·”œ«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«·≈”„"
GrdArray(5, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„ «·„ÿ«·»« "
oSearchDoc.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdNext_Click()
'openCardTable xdoc_no_zero.Caption, ">"
openCardTable xDoc_no.Text, ">"
If CardTable.EOF Then openCardTable xDoc_no.Text, "="
myload
End Sub
Private Sub CmdPrevious_Click()
'openCardTable xdoc_no_zero.Caption, "<"
openCardTable xDoc_no.Text, "<"
If CardTable.EOF Then openCardTable xDoc_no.Text, "="
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
Private Sub CmdNewInv_Click()
mydefine
On Error Resume Next
xCode.SetFocus
Err.Clear
End Sub

Private Sub CmdPrint_Click()
doprint
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
'openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
'openCardTable
myUndo
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdYearChange_Click()
If xType.BoundText = "2" And xDoc_no.Tag = LoadMode Then
    Years_LookupAll Me, oSearchYearChange
ElseIf xDoc_no.Tag = DefineMode Then
    Years_LookupAll Me, oSearchYearChange, "change_season"
End If
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    On Error Resume Next
    If sCode <> "" Then
        'xCode.Enabled = False
        xCode.Text = sCode
        xCode_LostFocus
        If bNew And sType <> "" Then
            mydefine
            xType.BoundText = sType
            cmdAddItems.SetFocus
            myAdditems
        End If
    Else
        If xDoc_no.Tag = LoadMode Then grid1.SetFocus Else xCode.SetFocus
    End If
    Err.Clear
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And ActiveControl.Name <> xCode.Name Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub
Private Sub Form_Load()
bEdit = True
bSlow = True

openCon con

Set DATA2.Recordset = myRecordSet("select * from paid_types", con)
Set xType.RowSource = DATA2
xType.ListField = "Desca"
xType.BoundColumn = "Code"

Set grid1.DataSource = DATA1
Set grid2.DataSource = DATA10


openCardTable
If bNew And sCode <> "" Then
    If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE6_20H.CODE = " & addvalue(sCode)
    mydefine
ElseIf sDoc_no <> "" Then
    myUndo
Else
    mydefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Set paidfrm = Nothing
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not MYVALID(True) Then
    On Error Resume Next
    grid1.SetFocus
    Err.Clear
    myloadgrd
    If Row < grid1.rows - 1 Then
        grid1.Select Row, Col
    Else
        CellPos 13, grid1.rows - 2, grid1.Cols - 1
    End If
    Exit Sub
End If

With grid1
If Not validRow(Row) Then Exit Sub
If Row = grid1.rows - 1 Then
    myAddItem
End If

On Error GoTo myerror
If MyReplace(Row) Then
    If xDoc_no.Tag = DefineMode Then
        Handlecontrols LoadMode
        myloadgrd
    ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
        myloadgrd
    End If
    CalcTotals
Else
    myloadgrd
End If
End With
Exit Sub
myerror:
myloadgrd
End Sub
Private Sub grid1_EnterCell()
With grid1
    If Not bEditRecord Then
        .Editable = flexEDKbdMouse
    ElseIf ((.Col = 0 And .TextMatrix(.Row, .Cols - 1) = "") Or .Col = 2 Or .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 6) Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xdate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not bIgMsg Then
    If grid1.rows < 3 Then
        MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
        Exit Function
    End If
End If

With grid1
For I = 1 To .rows - 2
'    If .TextMatrix(i, 1) = "" Then
'        .Select i, 0, i, grid1.Cols - 1
'        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
'        Exit Function
'    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
Dim I As Integer
xDoc_no.Text = CardTable!doc_no
'xdoc_no_zero.Caption = CardTable!doc_no_zero & ""
xForm_No.Text = CardTable!FORM_NO & ""
xdate.Text = myFormat_p(CardTable!Date)
xCode.Text = CardTable!code & ""
xYears_desca.Caption = CardTable!years_desca & ""

LoadMember

xType.BoundText = CardTable!Type & ""
xYears.Text = Myvalue(CardTable!years)
xYear_Desca.Caption = Year_Load(CardTable!YEAR_CODE, "DESCA_R", con)
xYear_code.Caption = CardTable!YEAR_CODE

xTotal_Year.Caption = Myvalue(CardTable!total_year)
xTotal_year_other.Caption = Myvalue(CardTable!total_year_other)
xTotal_Tax.Caption = Myvalue(CardTable!TOTAL_tax)
xTotal_late.Caption = Myvalue(CardTable!TOTAL_late)
xTotal.Caption = Myvalue(CardTable!total)

'xSeason.Caption = CardTable!SEASON
bClosed = True
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
bClosed = False

Handlecontrols LoadMode

myloadgrd2
myloadgrd

'CalcTotals

'cmd_closed.BackColor = IIf(CardTable!CLOSED, vbGreen, vbRed)
'cmd_closed.Caption = IIf(CardTable!CLOSED, "„€·ﬁ - › Õ «·„” ‰œ", "„› ÊÕ - ≈€·«ﬁ «·„” ‰œ")
'xusername.Caption = CardTable!UserName & ""
'xusername2.Caption = CardTable!UserName2 & ""
'XTIME1.Caption = Format(CardTable!Time, "YYYY-MM-DD HH:NN")
'xtime2.Caption = Format(CardTable!Time2, "YYYY-MM-DD HH:NN")

'CellPos index, 13, Grid1.rows - 2, Grid1.Cols - 1
On Error Resume Next
grid1.SetFocus
CellPos 13, grid1.rows - 2, grid1.Cols - 1
If grid2.rows > 1 Then grid2.Row = 1
Err.Clear
End Sub
Private Function myloadgrd(Optional bRefresh As Boolean = True) As Boolean
Dim cString As String
cString = "SELECT FILE6_20.[ITEM],FILE6_10.DESCA ,FILE6_20.[VALUE],[QUANT],[DISCOUNT_RATE],[TAX_RATE]" & _
          ",FILE6_20.LATE_RATE,FILE6_20.[TOTAL],FILE6_20.NOTES,FILE6_20.[ID]" & _
          " From [FILE6_20] INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM"
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_no.Text)
If grid2.rows < 2 Or grid2.Row < 1 Then
    cString = cString & turn(cString) & " FILE6_20.YEAR_CODE = " & addvalue(xYear_code.Caption)
Else
    If Not ValidNum(grid2.TextMatrix(grid2.Row, grid2.Cols - 1)) Then
        cString = cString & turn(cString) & " FILE6_20.YEAR_CODE = " & addvalue(xYear_code.Caption)
    Else
        cString = cString & turn(cString) & " FILE6_20.YEAR_CODE = " & grid2.TextMatrix(grid2.Row, grid2.Cols - 1)
    End If
End If
If bRefresh Then
    Set DATA1.Recordset = myRecordSet(cString, con)
    myAddItem
    Fixgrd
ElseIf DATA1.Recordset.Source <> cString Then
    Set DATA1.Recordset = myRecordSet(cString, con)
    myAddItem
    Fixgrd
    myloadgrd = True
End If

'cString = "SELECT FILE6_20.CODE,FILE1_20.DESCA,FILE6_20.MEMBER,FILE6_20.MEMBER_SUB,FILE6_20.DESCA,FILE6_20.VALUE,CONVERT(VARCHAR(10),FILE6_20.DATE_BIRTH,111),FILE6_20.NOTES,FILE6_20.[ID] " & _
'           " FROM FILE6_20 INNER JOIN FILE1_20 ON FILE6_20.CODE = FILE1_20.CODE " & _
'           " WHERE FILE6_20.Doc_no = " & MyParn(xDoc_No.Text)
'Data1.Refresh
'myAddItem
'End With
'Calctotals
End Function
Private Sub myloadgrd2()
Dim cString As String
cString = "SELECT YEARS_CODES.DESCA_R,YEARS_CODES.CODE " & _
          " From YEARS_CODES"
If Trim(xYears_desca.Caption) <> "" Then
    cString = cString & " WHERE CODE IN(" & xYears_desca.Caption & ")"
Else
    cString = cString & " WHERE CODE IS NULL"
End If
'cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & addvalue(xdoc_no.Text)
cString = cString & " ORDER BY CODE DESC"
Set DATA10.Recordset = myRecordSet(cString, con)
fixgrd2
End Sub
Private Sub mydefine()
Dim I As Integer, aRet As Variant
xDoc_no.Text = Newflag("FILE6_20H", "DOC_NO")

'xdoc_no_zero.Caption = ""
'xForm_no.Text = Newflag(cFileHeader, "FORM_NO", con, "SEASON = " & sSeason)
xForm_No.Text = ""
xType.BoundText = "1"

bClosed = True
xClosed.Value = 0
bClosed = False


xdate.Text = myFormat_p(Date)
aRet = Ret_Year(xdate.Text, , con)

xYear_Desca.Caption = retFlag(aRet, "desca")
xYear_code.Caption = retFlag(aRet, "code")
xYears_desca.Caption = ""

xCode.Text = sCode
xCodeDesca.Caption = ""
xType_Desca.Caption = ""
xUnPaid.Caption = ""
xLast_paid.Caption = ""
xType_Member.Caption = ""
If sCode <> "" Then LoadMember


xTotal_Year.Caption = ""
xTotal_year_other.Caption = ""
xTotal_Tax.Caption = ""
xTotal_late.Caption = ""
xTotal.Caption = ""

'cmd_closed.BackColor = &H8000000F
'cmd_closed.Caption = "-"
'xClosed.Value = 0
'xusername.Caption = ""
'xusername2.Caption = ""
'XTIME1.Caption = ""
'xtime2.Caption = ""

Fixgrd
grid1.rows = 1
myAddItem

fixgrd2
grid2.rows = 1

Handlecontrols DefineMode
CalcTotals
On Error Resume Next
'grid1.SetFocus
'Err.Clear
End Sub
Private Sub Handlecontrols(nMode)
bEditRecord = bEdit And xClosed.Value = 0 And xForm_No.Text = ""
cmdAddItems.Enabled = nMode = DefineMode
cmdYearChange.Enabled = nMode = LoadMode And xType.BoundText = "2"
'cmdFilter.Visible = cmdFilter.Tag <> ""
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = bEditRecord
cmddel.Enabled = nMode = LoadMode And bEditRecord
'xdate.Enabled = nMode = Mode
xdate.Locked = True
xForm_No.Locked = True
xCode.Locked = nMode = LoadMode
xType.Locked = nMode = LoadMode

'aRecords = retRecords(xdoc_no_zero.Caption)
aRecords = retRecords(xDoc_no.Text)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")

If nMode = LoadMode Then
    panel1(0).Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    panel1(0).Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If
cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sDoc_no = ""

xClosed.Enabled = nMode = LoadMode
xClosed.Enabled = xClosed.Enabled And (bopt1 Or xClosed.Value = 0)

xDoc_no.Enabled = (nMode = DefineMode)
xDoc_no.Tag = nMode
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not bEditRecord Then Exit Sub
With grid1
    If KeyCode = 112 And grid1.Col = 0 Then
        ItemsLookupAll Me, oSearchItems
    ElseIf KeyCode = 13 Then
        CellPos KeyCode, .Row, .Col
    ElseIf KeyCode = 46 And .Row <> .rows - 1 And .rows > 2 And bEditRecord Then
        If MsgBox("Õ–› øø", vbDefaultButton2 + vbOKCancel) = vbOK Then
            con.BeginTrans
            On Error GoTo myerror
            If .TextMatrix(.Row, .Cols - 1) <> "" Then
                con.Execute "Delete from FILE6_20 where ID = " & .TextMatrix(.Row, .Cols - 1)
            End If
            con.CommitTrans
            myRemove .Row
            grid1_EnterCell
        End If
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If (grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ €Ì— „”Ã·"
        Cancel = True
    ElseIf Not ValidInt(grid1.EditText) Then
        MsgBox "«·ﬂÊœ €Ì— ”·Ì„"
        Cancel = True
    Else
        If Not GrdDesc(grid1.EditText, Row) Then
           MsgBox "«·ﬂÊœ €Ì— ’ÕÌÕ «Ê ·« Ì’·Õ"
           Cancel = True
        End If
    End If
End If
End Sub

Private Sub grid2_Click()
If myloadgrd(False) Then
    CellPos 13, grid1.rows - 2, grid1.Cols - 1
End If
End Sub
Private Sub Grid2_EnterCell()
If myloadgrd(False) Then
    CellPos 13, grid1.rows - 2, grid1.Cols - 1
End If
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub xClosed_Click()
If Not bClosed Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.CLOSED = " & xClosed.Value & " WHERE DOC_NO = " & addvalue(xDoc_no.Text)
    con.CommitTrans
    Inform IIf(xClosed.Value = 1, " „ «€·«ﬁ «·„” ‰œ", " „ › Õ «·„” ‰œ")
    myUndo
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearchMember
ElseIf KeyCode = 13 And cmdAddItems.Enabled Then
    cmdAddItems_Click
End If
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
LoadMember
End Sub
Private Sub LoadMember()
xCodeDesca.Caption = ""
xType_Desca.Caption = ""
xLast_paid.Caption = ""
xUnPaid.Caption = ""
xDate_Begin.Caption = ""


If Not ValidInt(xCode.Text) Then Exit Sub
Dim aMember As Variant
aMember = Member_Load(xCode.Text, , con)
aPaid = Member_Paid(xCode.Text, , con)
If Not IsEmpty(aMember) Then
    xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
    xType_Desca.Caption = retFlag(aMember, "Type_Desca") & ""
    xType_Member.Caption = retFlag(aMember, "type") & ""
    xDate_Begin.Caption = myFormat_p(retFlag(aMember, "date_begin"))
End If

If Not IsEmpty(aPaid) Then
    xUnPaid.Caption = unpaid_years(retFlag(aPaid, "year_code"), sSeason, con)
    If retFlag(aPaid, "is_save") Then
        xLast_paid.Caption = "Õ«›Ÿ ⁄÷ÊÌ… Õ Ì " & retFlag(aPaid, "year_desca") & ""
    Else
        xLast_paid.Caption = "„”œœ Õ Ì " & retFlag(aPaid, "year_desca") & ""
    End If
Else
    xLast_paid.Caption = "·„ Ì”œœ „‰ ﬁ»·"
    xUnPaid.Caption = unpaid_years_count(xCode.Text, sSeason, con)
End If
'aUnPaid = retUnPaid(xCode.Text, sSeason, con, aPaid, aMember)
'xUnPaid.Caption = retFlag(aUnPaid, "Desca")
'xUnPaid_years.Caption = retFlag(aUnPaid, "Years")
End Sub
Private Sub xCurrent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
openCardTable
myUndo
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_no
If Not ValidNum(xDoc_no.Text) Then
     If xDoc_no.Tag = LoadMode Then
        mydefine
    Else
        xDoc_no.Text = Newflag("FILE6_20H", "DOC_NO", con)
    End If
Else
    If xDoc_no.Tag = LoadMode Then
        If (Not (CardTable.EOF)) Then
            If CardTable!doc_no = xDoc_no.Text Then
                Exit Sub
            End If
        End If
    End If
    
    openCardTable xDoc_no.Text
    If Not CardTable.EOF Then
        myload
    ElseIf xDoc_no.Tag = LoadMode Then
        mydefine
    Else
        xDoc_no.Text = Newflag("FILE6_20H", "DOC_NO", con)
    End If
End If
End Sub
Private Sub xForm_no_LostFocus()
Dim sDoc As String
If Trim(xForm_No.Text) = "" Then Exit Sub
'xDoc_No.Text = GetField("select top 1 doc_no from file6_20h where form_no = " & xForm_no.Text & " and season = " & xSeason.Caption)
'If xDoc_No.Text = "" Then xDoc_No.Text = GetField("select top 1 doc_no from file6_20h where form_no = " & xForm_no.Text)
'xDoc_No_LostFocus
End Sub
Private Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file2_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchItems.Caption = "≈” ⁄·«„ «·»‰Ê\"
oSearchItems.Show 1
End Sub
Private Function CalcTotals(Optional bOverRide As Boolean = False)
Dim nTotalRow As Double, Row As Integer
Dim Rate_Tax As Double, nTax As Double, nRate_Discount As Double, nDiscoun As Double, nRate_Late As Double, nLate As Double
'.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "‰”»… Œ’„|" & "‰”»… ﬁÌ„… „÷«›…|" & "‰”»… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"
With grid1
For Row = 1 To grid1.rows - 2
    nRate_Discount = mRound(.TextMatrix(Row, 4)) / 100
    nRate_tax = mRound(.TextMatrix(Row, 5)) / 100
    nRate_Late = mRound(.TextMatrix(Row, 6)) / 100
    
    nTotalRow = mRound(mRound(grid1.TextMatrix(Row, 2)) * mRound(grid1.TextMatrix(Row, 3)))
    nDiscount = mRound(mRound(nTotalRow) * nRate_Discount)
    
    nTotalRow = nTotalRow - nDiscount
    nTotalItem = nTotalItem + nTotalRow
    
    nTax = mRound(mRound(nTotalRow) * mRound(nRate_tax))
    nLate = mRound(mRound(nTotalRow) * mRound(nRate_Late))
    nTotalRow = nTotalRow + nTax + nLate
   .TextMatrix(Row, 7) = nTotalRow
    'nlate_total = mRound(Val(grid1.TextMatrix(Row, 9)), 2) + nlate_total
Next
If xDoc_no.Tag = LoadMode Then
    Dim aTotal As Variant
    aTotal = Doc_Totals(xDoc_no.Text, , con)
    
    xTotal_Year.Caption = Myvalue(retFlag(aTotal, "TOTAL_YEAR"))
    xTotal_year_other.Caption = Myvalue(retFlag(aTotal, "TOTAL_YEAR_OTHER"))
    xTotal_Tax.Caption = Myvalue(retFlag(aTotal, "TOTAL_TAX"))
    xTotal_late.Caption = Myvalue(retFlag(aTotal, "TOTAL_LATE"))
    xTotal.Caption = Myvalue(retFlag(aTotal, "TOTAL"))
End If
End With
'xTotal_items.Caption = nTotal_items
'xLate_Items.Caption = nLate_Items
'xLate_Total.Caption = Myvalue(nlate_total)
'xTotal.Caption = nTotal_items + nLate_Items + nlate_total

End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
'If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
.RowHeight(0) = 700
.FormatString = "ﬂÊœ «·»‰œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "⁄œœ|" & "‰”»… Œ’„|" & "‰”»… ﬁÌ„… „÷«›…|" & "‰”»… €—«„…|" & "«·≈Ã„«·Ì|" & "„·ÕÊŸ…|"
.ColWidth(0) = 800
.ColWidth(1) = 3000
.ColWidth(2) = 1100
.ColWidth(3) = 1000
.ColWidth(4) = 1100
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1200
.ColWidth(8) = 3000
.ColWidth(9) = 1000
.ColDataType(3) = flexDTDecimal
.ColDataType(4) = flexDTDecimal
.ColDataType(5) = flexDTDecimal
.ColDataType(6) = flexDTDecimal
.ColDataType(7) = flexDTDecimal
.ColDataType(8) = flexDTDecimal




'.ColHidden(4) = True
'.ColHidden(6) = True
'.ColHidden(7) = True
'.ColHidden(9) = True

.ColHidden(.Cols - 1) = True
For I = 1 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColComboList(0) = cList
End With
End Sub
Private Sub fixgrd2()
With grid2
.FormatString = "«·”‰…|" & "«·ﬂÊœ"
.ColWidth(0) = 2000
'.ColWidth(1) = 1200
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Fixgrd3()
With Grid3
'.FormatString = "«·”‰…|" & "«·≈Ã„«·Ì|" & "«·ﬂÊœ"
'.ColWidth(0) = 2000
'.ColWidth(1) = 1200
'.ColHidden(.Cols - 1) = True
'For i = 0 To .Cols - 1
'    .ColAlignment(i) = flexAlignRightCenter
'Next
End With
End Sub
Private Function openCardTable(Optional pDoc_no As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String

Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1  FILE6_20H.* FROM FILE6_20H"

If pSign = "=" Then
    If pDoc_no <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_no)
Else
    If pDoc_no <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_no)
    'If pDoc_No <> "" Then cWhere = "DOC_NO_ZERO " & pSign & addstring(pDoc_No)
End If

'If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addvalue(pDoc_No)

cFilter = ""
cFilter = "FILE6_20H.OLD = 0"

If xCurrent.Value = 1 Then
    Dim aRet As Variant
    'aRet = Year_Load(sSeason, , con)
    'cFilter = "FILE6_20H.DATE >= " & DateSq(retFlag(aRet, "DATE1"))
    'cFilter = cFilter & " AND FILE6_20H.DATE <= " & DateSq(retFlag(aRet, "DATE2"))
    cFilter = "FILE6_20H.YEAR_CODE = " & sSeason
End If

'If cmdFilter.Tag <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE6_20H.DOC_NO IN(" & cmdFilter.Tag & ")"
'If chkDay.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "FILE6_20H.[DATE] = " & DateSq(retDate)
'If chkMonth.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & Year(Date) & " AND MONTH(DATE) = " & Month(Date)
'If chkYear.Value = 1 Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & Year(Date)
'If cmdYear.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "YEAR(FILE6_20H.[DATE]) = " & cmdYear.Tag
'If cmdsup.Tag <> "" Then cFilter = cFilter & turn(cFilter, " And ") & "FILE6_20H.CODE  = " & MyParn(oSearchCode.grid1.TextMatrix(oSearchCode.grid1.Row, 0)) & ")"

'If sCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "FILE6_20H.CODE = " & addvalue(sCode)

If sDoc_no <> "" Then cFilter = "FILE6_20H.DOC_NO = " & addvalue(sDoc_no)

If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
     cString = cString & " order by FILE6_20H.doc_no desc"
    'cString = cString & " order by FILE6_20H.doc_no_zero desc"
ElseIf pSign = ">=" Or pSign = ">" Then
     cString = cString & " order by FILE6_20H.doc_no ASC"
    'cString = cString & " order by FILE6_20H.doc_no_zero ASC"
End If
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Function retRecords(pDoc_no) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_no <> "" Then
    'cString = "SELECT SUM(1) AS records,SUM(CASE WHEN doc_no_zero <= " & MyParn(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN doc_no <= " & addvalue(pDoc_no) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM file6_20H " & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If pDoc_no <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String
If ValidNum(xDoc_no.Text) Then
    openCardTable xDoc_no.Text
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
Private Sub myAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        myRemove OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
'If (Not validRow(grid1.Row)) And grid1.Row <> grid1.rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then myRemove grid1.Row
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
With grid1
If Col < .Cols - 7 Then
    .Col = Col + 1 + IIf(Col = 0, 2, 0)
ElseIf Row < .rows - 1 Then
    .Select Row + 1, NextEmpty(grid1, Row + 1, 0, 3)
    .ShowCell .Row, 0
Else
    .Select Row, Col
End If
End With
End Sub
Private Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
For I = IIf(nBegincol = -1, pGrid.Cols - 1, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
    If Trim(pGrid.TextMatrix(Row, I)) = "" Then
        NextEmpty = I
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function

Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
CalcTotals
End Sub
Private Function GrdDesc(sItem As String, Row As Long) As Boolean
Dim sSection As String, aRet As Variant
If ValidNum(sItem) Then
    If ValidNum(xCode.Text) And RetYearCode Then
        aRet = GetFields("SELECT TOP 1 DESCA,dbo.f_mem_item_value(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS VALUE,dbo.f_mem_item_discount(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS DISCOUNT,dbo.f_mem_item_tax(" & sItem & "," & xCode.Text & "," & RetYearCode & ") AS TAX   FROM file6_10 where ITEM = " & sItem)
    Else
        aRet = GetFields("SELECT TOP 1 DESCA,VALUE FROM file6_10 where ITEM = " & sItem)
    End If
    grid1.TextMatrix(Row, 1) = retFlag(aRet, "DESCA") & ""
    grid1.TextMatrix(Row, 2) = retFlag(aRet, "VALUE") & ""
    If grid1.TextMatrix(Row, 3) = "" Then grid1.TextMatrix(Row, 3) = 1
    grid1.TextMatrix(Row, 4) = mRound(retFlag(aRet, "DISCOUNT"))
    grid1.TextMatrix(Row, 5) = mRound(retFlag(aRet, "TAX"))
End If
GrdDesc = True
End Function
Private Function doprint()
If Not MYVALID Then Exit Function
Dim loctable As ADODB.Recordset, cString As String
Dim temptable As New ADODB.Recordset
cString = "SELECT FILE6_20.DOC_NO,TYPE_CODES.DESCA AS TYPE_DESCA,FILE6_20H.CODE,FILE6_20H.DATE,FILE6_20H.TOTAL_YEAR,FILE6_20H.TOTAL_YEAR_OTHER,FILE6_20.TOTAL,FILE1_10.DESCA AS DESCA_MEMBER,FILE6_10.DESCA AS ITEM_DESCA," & _
          "(FILE6_20.TOTAL_ITEM),FILE6_20.TOTAL AS TOTAL_ROW,FILE6_20.TAX,FILE6_20.TOTAL_DISCOUNT ,FILE6_20.DISCOUNT_RATE,FILE6_20H.TOTAL_LATE,FILE6_20H.TOTAL_TAX,FILE6_20H.TOTAL" & _
          " FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO " & _
          " INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE" & _
          " INNER JOIN TYPE_CODES ON FILE1_10.TYPE = TYPE_CODES.CODE" & _
          " INNER JOIN FILE6_10 ON FILE6_20.ITEM = FILE6_10.ITEM" & _
          " WHERE FILE6_20.YEAR_CODE = FILE6_20H.YEAR_CODE"
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & xDoc_no.Text


Dim aTotal As Variant
'aTotal = GetFields("Select sum(file6_20.total) as total from file6_20 where doc_no = " & xDoc_No.Text)
Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

Dim I As Long
With loctable
Do Until loctable.EOF
    temptable.AddNew
    temptable!str1 = ArbString(myFormat_p(loctable!Date))
    temptable!str2 = loctable!Desca_Member
    temptable!Str3 = ArbString(loctable!code)
    temptable!str4 = loctable!TYPE_DESCA
    temptable!str5 = ArbString(loctable!doc_no)
    temptable!val1 = mRound(loctable!total_item)
    temptable!val2 = mRound(loctable!discount_Rate)
    'temptable!Val3 = mRound(loctable!total_discount)
    temptable!Val3 = mRound(loctable!total_row)
        
    temptable!STR6 = loctable!item_Desca
    temptable!val4 = mRound(loctable!total_year)
    temptable!Val5 = mRound(loctable!total_year_other) + mRound(loctable!TOTAL_late)
    temptable!Val6 = 0
    temptable!Val7 = mRound(loctable!TOTAL_tax)
    
    temptable!Val8 = mRound(loctable!total)
    temptable!Val15 = mRound(loctable!tax)
    
    nTotal = mRound(loctable!total)
    
    'temptable!str11 = TurnValue(loctable!Item_Desca)
    'temptable!str12 = TurnValue(loctable!Desca)
    'temptable!str13 = TurnValue(loctable!notes)
    'temptable!str13 = TurnValue(loctable!notes)
    'temptable!str14 = TurnValue(loctable!user_name)
    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
    'temptable!VAL1 = Val(loctable! & "")
    temptable!Str10 = loctable!total
    temptable!Val10 = I
    temptable!VAL20 = 1
    temptable.Update
    loctable.MoveNext
Loop

Set loctable = Nothing

cString = "SELECT FILE1_11.DESCA,FILE1_11.DATE_BIRTH ,RELATION_CODES.DESCA AS RELATION_DESCA " & _
          " FROM FILE1_11 INNER JOIN RELATION_CODES ON FILE1_11.RELATION = RELATION_CODES.CODE" & _
          " WHERE FILE1_11.MEMBER = " & addvalue(xCode.Text)
pDate = Year_Load(xYear_code.Caption, "DATE2", con)
If xType.MatchedWithList Then
    Dim atype As Variant
    atype = Claim_Type_Load(xType.BoundText, , con)
    If retFlag(atype, "over_age") Then
        cString = cString & " AND (NOT(GENDER = 1 AND HANDI = 0 AND RELATION = 2 AND dbo.f_age(FILE1_11.DATE_BIRTH ," & addstring(myFormat(pDate)) & ") > 24))"
    End If
End If
cString = cString & " ORDER BY FILE1_11.RELATION,FILE1_11.DATE_BIRTH"

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

Do Until loctable.EOF
    temptable.AddNew
    temptable!str1 = loctable!desca
    temptable!str2 = TurnValue(ArbString(myFormat_p(loctable!DATE_BIRTH)))
    temptable!Str3 = loctable!RELATION_DESCA
    temptable!STR15 = ArbString(MyOnly(mRound(nTotal)))
    temptable!str21 = "≈Ì’«· ”œ«œ Ê«” ·«„"
    'temptable!VAL1 = Val(loctable! & "")
    temptable!Val10 = I
    temptable!VAL20 = 2
    temptable.Update
    loctable.MoveNext
Loop

End With
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
REPORT1.WindowState = crptMaximized
REPORT1.ReportFileName = App.Path & "\Reports\paid.rpt"
REPORT1.DataFiles(0) = tempFile
REPORT1.WindowShowPrintSetupBtn = True
REPORT1.ProgressDialog = False
REPORT1.CopiesToPrinter = 1

iSubreports = REPORT1.GetNSubreports
If (iSubreports <> 0) Then
    For I = 0 To iSubreports - 1
        sSubreportName = REPORT1.GetNthSubreportName(I)
        REPORT1.SubreportToChange = sSubreportName
        REPORT1.DataFiles(0) = tempFile
    Next
End If
'REPORT1.Destination = crptToPrinter
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Function

Private Function addPaymentYears()

End Function
Public Function addPaidItems2(pDoc_no As String, pYear As Integer, pType, aMember As Variant, pCon As ADODB.Connection) As String
'Dim cString As String, nAge As Long, bMemberAdd As Boolean
'Dim nAll As Long, aPen(3) As Integer
'
'cString = "SELECT FILE6_11.ITEM,FILE6_10.AGE1,FILE6_10.AGE2 ,FILE6_10.DESCA, FILE6_10.ALLMEMBER, FILE6_10.LATE, FILE6_10.RELATION," & _
'      " FILE6_10.ISMEMBER, COALESCE(FILE6_10.AGE1,0), COALESCE(FILE6_10.AGE2,0), FILE6_10.GENDER, " & _
'      " FILE6_10.BASICDIED, FILE6_10.BASICNEW,FILE6_10.BASICOLD, FILE6_10.MEETING, FILE6_10.DAYS, FILE6_10.NORATE, " & _
'      " FILE6_11.value, FILE6_11.Discount " & _
'      " FROM FILE6_10 INNER JOIN FILE6_11 ON FILE6_10.ITEM = FILE6_11.item " & _
'      " WHERE FILE6_11.TYPE = " & pType & _
'      " AND FILE6_11.BASIC = 1 " & _
'      " AND FILE6_11.YEAR_CODE = " & pYear & _
'      " AND [SECTION] =  " & retFlag(aMember, "type")
'cString = cString & " ORDER BY FILE6_10.ITEM"
'
'Dim loctable As ADODB.Recordset
'Set loctable = New ADODB.Recordset
'loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'
'bMemberAdd = retFlag(aMember, "Died", False)
'Do Until loctable.EOF
'    If loctable!isMember Then
'        If (Not bMemberAdd) Then
'            If AddMemberData(loctable, aMember, Index, pDate) Then
'                aInsert = AddFlag(Empty, "doc_no", pDoc_No)
'                aInsert = AddFlag(aInsert, "item", loctable!Item)
'                aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
'                aInsert = AddFlag(aInsert, "quant", "1")
'                aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
'                If loctable!late Then
'                    aInsert = AddFlag(aInsert, "late_rate", aPen(Index))
'                End If
'                aInsert = AddFlag(aInsert, "TAB", Index)
'                cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
'                bMemberAdd = True
'            End If
'        End If
'    ElseIf Not IsNull(loctable!RELATION) Then
'        nRelation = addRelation(loctable, loctable!RELATION, retFlag(aMember, "code") & "", pDate, pCon)
'        If nRelation > 0 Then
'            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
'            aInsert = AddFlag(aInsert, "item", loctable!Item)
'            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
'            aInsert = AddFlag(aInsert, "quant", nRelation)
'            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
'            If loctable!late Then
'                aInsert = AddFlag(aInsert, "late_rate", aPen(Index))
'            End If
'            aInsert = AddFlag(aInsert, "TAB", Index)
'            cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
'        End If
'    ElseIf loctable!BasicNew Or loctable!basicOld Then
'        If (loctable!BasicNew And IsEmpty(aPaid)) Or (loctable!basicOld And Not IsEmpty(aPaid)) Then
'            aInsert = AddFlag(Empty, "doc_no", pDoc_No)
'            aInsert = AddFlag(aInsert, "item", loctable!Item)
'            aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
'            aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
'            aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
'            If loctable!late Then
'                aInsert = AddFlag(aInsert, "late_rate", aPen(Index))
'            End If
'            aInsert = AddFlag(aInsert, "TAB", Index)
'            cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
'        End If
'    Else
'        aInsert = AddFlag(Empty, "doc_no", pDoc_No)
'        aInsert = AddFlag(aInsert, "item", loctable!Item)
'        aInsert = AddFlag(aInsert, "value", Val(loctable!Value))
'        aInsert = AddFlag(aInsert, "quant", IIf(loctable!AllMember, nAll, 1))
'        aInsert = AddFlag(aInsert, "discount_rate", Val(loctable!discount & ""))
'        If loctable!late Then
'            aInsert = AddFlag(aInsert, "late_rate", aPen(Index))
'        End If
'        aInsert = AddFlag(aInsert, "TAB", Index)
'        cInsert = cInsert & addInsert(aInsert, "FILE6_20") & ";"
'    End If
'    loctable.MoveNext
'Loop
'addPaidItems = cInsert
End Function
Private Function RetYearCode() As String
If grid2.Row < 0 Or grid2.rows < 1 Then
    RetYearCode = xYear_code.Caption
ElseIf Not ValidNum(grid2.TextMatrix(grid2.Row, grid2.Cols - 1)) Then
    RetYearCode = xYear_code.Caption
Else
    RetYearCode = grid2.TextMatrix(grid2.Row, grid2.Cols - 1)
End If
End Function
Private Sub updateYears(pDoc_no)
'con.Execute "update file6_20h set years_desca = dbo.f_get_years(" & pDoc_no & ") where file6_20h.doc_no = " & pDoc_no & ";"
End Sub
Private Sub xYears_GotFocus()
myGotFocus xYears
End Sub
Private Sub xYears_LostFocus()
myLostFocus xYears
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDate_GotFocus()
myGotFocus xdate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xdate
myValidDate xdate
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_no
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Function ChangeYear(pYear) As Boolean
Dim aInsert As Variant, I As Integer
aInsert = AddFlag(Empty, "[year_code]", pYear)
con.BeginTrans
On Error GoTo myerror
con.Execute addUpdate(aInsert, "FILE6_20H", "doc_no = " & addstring(xDoc_no.Text))
con.CommitTrans
ChangeYear = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
