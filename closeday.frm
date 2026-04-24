VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form closedayfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "d"
   ClientHeight    =   10635
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
   ScaleHeight     =   10635
   ScaleWidth      =   18060
   WindowState     =   2  'Maximized
   Begin Olymbic.CSubclass CSubclass1 
      Left            =   -585
      Top             =   -360
      _ExtentX        =   2487
      _ExtentY        =   1296
   End
   Begin VB.CheckBox xDay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·ÌÊ„ ›ﬁÿ"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   14985
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   7650
      Width           =   1275
   End
   Begin VB.CheckBox xCurrent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   ".«·„Ê”„ «·Õ«·Ì ›ﬁÿ"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   16875
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   7650
      Value           =   1  'Checked
      Width           =   1815
   End
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
      Left            =   14940
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   3795
      TabIndex        =   41
      Top             =   5940
      Width           =   3795
      Begin Threed.SSCommand cmdFirst 
         Height          =   420
         Left            =   2880
         TabIndex        =   42
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
         Picture         =   "closeday.frx":0000
         Caption         =   "√Ê·"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "closeday.frx":21A7
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   1890
         TabIndex        =   43
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
         Picture         =   "closeday.frx":41EE
         Caption         =   "”«»ﬁ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "closeday.frx":62D9
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   990
         TabIndex        =   44
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
         Picture         =   "closeday.frx":82D3
         Caption         =   " «·Ì"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "closeday.frx":A3E4
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   45
         TabIndex        =   45
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
         Picture         =   "closeday.frx":C3DE
         Caption         =   "«ŒÌ—"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "closeday.frx":E602
      End
   End
   Begin VB.TextBox xDoc_no 
      Alignment       =   1  'Right Justify
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
      Left            =   12510
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Tag             =   "2"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame fmBox 
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
      Height          =   645
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   6885
      Width           =   3795
      Begin Threed.SSCommand cmdUndo 
         Height          =   465
         Left            =   45
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   820
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
         Picture         =   "closeday.frx":106D3
         Caption         =   "√⁄«œ…  «·÷»ÿ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "closeday.frx":1284A
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1005
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   4905
      Width           =   3750
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì €Ì— „”œœ…"
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
         Index           =   0
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   585
         Width           =   2010
      End
      Begin VB.Label xtotal_unpaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ «Ì’«·«  €Ì— „”œœ…"
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
         Index           =   3
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1875
      End
      Begin VB.Label xInv_count_unpaid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   1320
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
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3150
      Width           =   3750
      Begin VB.Label xtotal_net 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "’«›Ì «·«Ì—«œ"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1305
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì ﬁÌ„… „÷«›…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   945
         Width           =   1650
      End
      Begin VB.Label xTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "≈Ã„«·Ì «·«Ì’«·«  «·„”œœ…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label xtotal_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label xInv_count_paid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "⁄œœ «·«Ì’«·«  «·„”œœ…"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1830
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   10485
      Visible         =   0   'False
      Width           =   6585
      Begin VB.Frame frmDrawn 
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
         Height          =   1095
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1350
         Width           =   3525
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·„”ÕÊ»«  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   630
            Width           =   1155
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "»«ﬁÌ :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   315
            Width           =   1065
         End
      End
      Begin MSDataListLib.DataCombo xBox3 
         Height          =   330
         Left            =   180
         TabIndex        =   32
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label xMinCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label xPlus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2475
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label xReserve 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2115
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "”œ«œ ¬Ã· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2475
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblTrans 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÕÃ“ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2115
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„— Ã⁄«  ⁄—»Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label xReturns 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Œœ„… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   765
         Width           =   1650
      End
      Begin VB.Label Label13 
         Caption         =   "»Ì⁄ √Ã· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Œ“‰… «·‘—ﬂ«¡ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Œ“‰… «·»Ì⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   3075
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   3705
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   180
         Width           =   2085
      End
      Begin MSACAL.Calendar xdate_cal 
         Height          =   2430
         Left            =   90
         TabIndex        =   49
         Top             =   585
         Width           =   3630
         _Version        =   524288
         _ExtentX        =   6403
         _ExtentY        =   4286
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2006
         Month           =   5
         Day             =   21
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483635
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   7965
      Top             =   9900
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
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   3450
      Left            =   10035
      TabIndex        =   10
      Top             =   9900
      Visible         =   0   'False
      Width           =   2010
      _cx             =   3545
      _cy             =   6085
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   0   'False
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   12.0098039215686
      ZoomMode        =   4
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   90
      TabIndex        =   22
      Top             =   1125
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "ﬂ· «·«” „«—« "
      TabPicture(0)   =   "closeday.frx":14B37
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«” „«—«  „”Ã·… Ê€Ì— „”Ã·…"
      TabPicture(1)   =   "closeday.frx":14B53
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(1)=   "grid2"
      Tab(1).Control(2)=   "data1"
      Tab(1).Control(3)=   "data13"
      Tab(1).Control(4)=   "DATA11"
      Tab(1).Control(5)=   "DATA12"
      Tab(1).ControlCount=   6
      Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   4785
         Left            =   -74910
         TabIndex        =   23
         Top             =   360
         Width           =   13965
         _cx             =   24633
         _cy             =   8440
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid Grid3 
         Height          =   8205
         Left            =   45
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   360
         Width           =   14550
         _cx             =   25665
         _cy             =   14473
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
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4200
         Left            =   -74955
         TabIndex        =   0
         Top             =   405
         Width           =   14595
         _cx             =   25744
         _cy             =   7408
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
         TabBehavior     =   0
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   3885
         Left            =   -74955
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4635
         Width           =   14595
         _cx             =   25744
         _cy             =   6853
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
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   -74955
         Top             =   630
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
      Begin MSAdodcLib.Adodc data13 
         Height          =   330
         Left            =   -73065
         Top             =   2160
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
      Begin MSAdodcLib.Adodc DATA11 
         Height          =   330
         Left            =   -70275
         Top             =   3870
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
      Begin MSAdodcLib.Adodc DATA12 
         Height          =   330
         Left            =   -68925
         Top             =   3285
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
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   1395
      TabIndex        =   38
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
      Picture         =   "closeday.frx":14B6F
      Caption         =   "Œ—ÊÃ"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   690
      Left            =   5265
      TabIndex        =   39
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
      Picture         =   "closeday.frx":16E92
      Caption         =   " ﬁ—Ì— «·ÌÊ„"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "closeday.frx":19208
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   45
      Width           =   6585
      Begin VB.TextBox xForm_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4095
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Tag             =   "N"
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox xDoc_Pc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   135
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "N"
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4095
         MaxLength       =   9
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «Ì’«· «·”œ«œ"
         Height          =   240
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   675
         Width           =   990
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·„” ‰œ"
         Height          =   240
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         Height          =   240
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   315
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Width           =   3930
      End
   End
   Begin Threed.SSCommand cmdGrdItem 
      Height          =   690
      Left            =   3060
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   405
      Width           =   2175
      _ExtentX        =   3836
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
      Picture         =   "closeday.frx":1B38B
      Caption         =   " ﬁ—Ì— «·”œ«œ «·ÌÊ„Ì"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      PictureDisabled =   "closeday.frx":1D701
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   465
      Left            =   990
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6165
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
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
      Picture         =   "closeday.frx":1F884
      Caption         =   "√⁄«œ…  «·÷»ÿ"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "closeday.frx":219FB
   End
   Begin Threed.SSCommand cmdDayBefore 
      Height          =   510
      Left            =   14985
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   8640
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   900
      _Version        =   196610
      CaptionStyle    =   1
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
      Caption         =   "”Õ» «Ì’«·«  ”«»ﬁ…"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "closeday.frx":23CE8
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   14985
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   7920
      Width           =   3705
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "⁄÷ÊÌ… „ﬁ”ÿ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   270
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "⁄÷ÊÌ… ⁄«„·…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   270
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Label xRecord_No 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   14940
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   6525
      Width           =   3750
   End
End
Attribute VB_Name = "closedayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aData As Variant
Dim sFile As String, bAct As Boolean
Dim CardTable As ADODB.Recordset
Public bIgPend As Boolean, bCancel As Boolean, oSearchDoc As New Search
Public sDate As String, Sbox As String, myForm As Form, bEnd  As Boolean, cFilter As String
Dim oSearch As New Search
Dim con As New ADODB.Connection
Const LoadMode = 1, DefineMode = 2
Private Sub CmdDel_Click()
If xClosed.Value = 0 Then
    con.BeginTrans
    If Option1(0).Value Then
        con.Execute "delete from closeday where code = " & xCode.text
    Else
        con.Execute "delete from closeday_install where code = " & xCode.text
    End If
    con.CommitTrans
    On Error Resume Next
    Unload menufrm
    Unload Me
End If
End Sub
Private Sub cmdDay_Click()
xDate.text = myFormat(Date)
xdate_LostFocus
End Sub
Private Sub cmdDayBefore_Click()

If Option1(0).Value Then
   DocLookup
Else
   DocLookup_install
End If
'Dim cString As String, pDate As String, nRecords As Long
'If Option1(0).Value Then
'    cString = "select top 1 DATE from v_close_day where date < " & DateSq(xDate.Text) & "  order by date desc"
'Else
'    cString = "select top 1 DATE from v_close_day_install where date < " & DateSq(xDate.Text) & "  order by date desc"
'End If
'
'pDate = myFormat(GetField(cString, con))
'If IsDate(pDate) Then
'    con.BeginTrans
'    On Error GoTo myerror
'    If Option1(0).Value Then
'        con.Execute "UPDATE FILE6_20H SET DATE = " & DateSq(xDate.Text) & " WHERE DATE_ISSUE = " & DateSq(pDate), nRecords
'    Else
'        con.Execute "UPDATE FILE6_30H SET DATE = " & DateSq(xDate.Text) & " WHERE DATE_ISSUE = " & DateSq(pDate), nRecords
'    End If
'    con.CommitTrans
'    MsgBox " „  ÕÊÌ· " & nRecord & " «” „«—…"
'    myUndo
'End If

Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub DocLookup()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_20H.DOC_NO,PAID_TYPES.DESCA,CONVERT(VARCHAR(10),FILE6_20H.DATE,111),YEARS_CODES.DESCA,FILE6_20H.CODE, FILE1_10.DESCA" & _
          "  FROM  FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE INNER JOIN YEARS_CODES ON FILE6_20H.YEAR_CODE = YEARS_CODES.CODE"
cString = cString & " WHERE FILE6_20H.OLD = 0"
cString = cString & " AND FILE6_20H.YEAR_CODE >= " & sSeason
cString = cString & " AND FILE6_20H.DATE < " & DateSq(xDate.text)

'cString = cString & " AND FILE6_20H.DATE > " & DateSq(DateAdd("d", -14, xDate.Text))
'cString = cString & " AND FILE6_20H.DATE <> " & DateSq(xDate.Text)
cString = cString & " AND FILE6_20H.FORM_NO IS NULL"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_20H.DATE DESC,FILE6_20H.YEAR_CODE DESC,FILE6_20H.Doc_No DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
listarray(0, 1) = "(**FILE6_20H.CODE**)"

listarray(1, 0) = "—ﬁ„ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(1, 1) = "(%%FILE1_10.Desca%% or **FILE6_20H.DOC_NO**)"

listarray(2, 0) = "‰Ê⁄ «·„ÿ«·»…"
listarray(2, 1) = "(**FILE6_20H.[TYPE]**)"
listarray(2, 2) = "SELECT CODE,DESCA FROM PAID_TYPES"
listarray(2, 3) = "CODE"
listarray(2, 4) = "DESCA"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1400

GrdArray(1, 0) = "‰Ê⁄ «·„” ‰œ"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(2, 1) = 1350

GrdArray(3, 0) = "”‰… «·”œ«œ"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
GrdArray(4, 1) = 1300

GrdArray(5, 0) = "«·≈”„"
GrdArray(5, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = "«” ⁄·«„ «·„ÿ«·»« "
oSearchDoc.bUnLoad = True
oSearchDoc.Show 1
End Sub
Private Sub DocLookup_install()
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT TOP 2000 FILE6_30H.DOC_NO,CONVERT(VARCHAR(10),FILE6_30H.DATE,111),FILE2_10.CODE, FILE2_10.DESCA" & _
          "  FROM  FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE"
'cString = cString & " AND FILE6_30H.DATE > " & DateSq(DateAdd("d", -30, xDate.Text))
cString = cString & " AND FILE6_30H.DATE <> " & DateSq(xDate.text)
cString = cString & " AND FILE6_30H.FORM_NO IS NULL"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter

Generalarray(1) = cString
Generalarray(2) = " ORDER BY FILE6_30H.DATE DESC,FILE6_30H.Doc_No DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
listarray(0, 1) = "(**FILE6_30H.CODE**)"

listarray(1, 0) = "—ﬁ„ «·„” ‰œ-«”„ «·⁄÷Ê"
listarray(1, 1) = "(%%FILE2_10.Desca%% or **FILE6_30H.DOC_NO**)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1400

GrdArray(2, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sCaption = "«” ⁄·«„ „ÿ«·»«   ﬁ”Ìÿ"
oSearchDoc.bUnLoad = True
oSearchDoc.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
If Option1(0).Value And ValidNum(xCode.text) Then
    aPaid = Member_Paid(xCode.text, , con)
    If retFlag(aPaid, "YEAR_CODE") = sSeason Then
        MsgBox ArbString("«·⁄÷Ê „”œœ » «—ÌŒ : " & myFormat_p(retFlag(aPaid, "date")) & vbCrLf & "«Ì’«· —ﬁ„ : " & retFlag(aPaid, "form_no")), vbCritical
    End If
End If
myLoadGrd1
If grid1.rows > 1 Then
    CellPos 13, 1, grid1.Cols - 1
    On Error Resume Next
    Err.Clear
End If
myloadgrd2
Myloadgrd3
End Sub

Private Sub cmdClose_Click()
End Sub
Private Sub cmdGrdItem_Click()
If Option1(0).Value Then
    grditemGroupfrm.sDate1 = xDate.text
    grditemGroupfrm.sdate2 = xDate.text
    grditemGroupfrm.Show
Else
    grditemGroup_installfrm.sDate1 = xDate.text
    grditemGroup_installfrm.sdate2 = xDate.text
    grditemGroup_installfrm.Show
End If
End Sub

Private Sub CmdUndo_Click()
myUndo
End Sub

Private Sub Command1_Click()
xCode.SetFocus
End Sub

Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    xCode.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Name <> xCode.Name And ActiveControl.Name <> xForm_No.Name And ActiveControl.Name <> xDoc_Pc.Name Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
CSubclass1.SubClassMe SSTab1.hwnd, 0, , vbWhite       '//--- Begin SubClassing
Err.Clear

openCon con
'CSubclass1.SubClassMe Me.SSTab1.hWnd, 0, , vbWhite      '//--- Begin SubClassing
'aSession = GetFields("Select * from closeday where date = " & DateSq(dSalesDate) & " and box1 = " & MyParn(retFlag(aSec, "Box_sales")))
'If Not IsDate(xDate.text) Then xDate.text = Format(GetDate, "yyyy/mm/dd")
'aData = GetFields("select desca,address,phone,mail from address")


Set grid1.DataSource = DATA11
Set grid2.DataSource = DATA12
Set grid3.DataSource = DATA13


fixgrd1
fixgrd2
Fixgrd3
'fixGrd4

openCardTable
xDate.text = myFormat(Date)
xdate_LostFocus

'myUndo
'openCardTable
End Sub
Private Sub myLoadGrd1()
Dim cString As String
If Option1(0).Value Then
    cString = "SELECT  FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,FILE6_20H.CODE,FILE1_10.DESCA,FILE6_20H.TOTAL_YEAR,FILE6_20H.TOTAL_YEAR_OTHER,FILE6_20H.TOTAL_TAX,FILE6_20H.TOTAL_LATE,FILE6_20H.TOTAL,FILE6_20H.YEAR_CODE" & _
              " FROM FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "(FILE6_20H.FORM_NO IS NULL)"
    cString = cString & turn(cString) & "FILE6_20H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_20H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_20H.CODE = " & addvalue(xCode.text)
    cString = cString & " ORDER by FILE6_20H.DOC_NO"
    Set DATA11.Recordset = myRecordSet(cString, con)
    fixgrd1
Else
    cString = "SELECT FILE6_30H.DOC_NO,FILE6_30H.FORM_NO,FILE6_30H.FORM_NO2,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.TOTAL_VALUE,FILE6_30H.CARD_VALUE,FILE6_30H.TOTAL_TAX,FILE6_30H.INTEREST,FILE6_30H.OTHER,FILE6_30H.CHARGE,FILE6_30H.TOTAL" & _
              " FROM FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE WHERE IsFawry = 0"
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_30H.DOC_NO = " & addvalue(xDoc_Pc.text)
    cString = cString & turn(cString) & "((FILE6_30H.FORM_NO IS NULL) OR (FILE6_30H.FORM_NO2 IS NULL))"
    cString = cString & turn(cString) & "FILE6_30H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_30H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_30H.CODE = " & addvalue(xCode.text)
    cString = cString & " ORDER by FILE6_30H.DOC_NO"
    Set DATA11.Recordset = myRecordSet(cString, con)
    fixgrd1_install
End If
End Sub
Private Sub myloadgrd2()
Dim cString As String
If Option1(0).Value Then
    cString = "SELECT FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,FILE6_20H.CODE,FILE1_10.DESCA,FILE6_20H.TOTAL_YEAR,FILE6_20H.TOTAL_YEAR_OTHER,FILE6_20H.TOTAL_TAX,FILE6_20H.TOTAL_LATE,FILE6_20H.TOTAL,FILE6_20H.YEAR_CODE" & _
              " FROM FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "(NOT FILE6_20H.FORM_NO IS NULL)"
    cString = cString & turn(cString) & "FILE6_20H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_20H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_20H.CODE = " & addvalue(xCode.text)
    If ValidNum(xForm_No.text) Then
        cString = cString & turn(cString) & "FILE6_20H.FORM_NO = " & DateSq(xForm_No.text)
    End If
    cString = cString & " ORDER by FILE6_20H.FORM_NO"
    Set DATA12.Recordset = myRecordSet(cString, con)
    fixgrd2
Else
    cString = "SELECT FILE6_30H.DOC_NO,FILE6_30H.FORM_NO,FILE6_30H.FORM_NO2,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.TOTAL_VALUE,FILE6_30H.CARD_VALUE,FILE6_30H.TOTAL_TAX,FILE6_30H.INTEREST,FILE6_30H.OTHER,FILE6_30H.CHARGE,FILE6_30H.TOTAL" & _
              " FROM FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "((NOT FILE6_30H.FORM_NO IS NULL) AND (NOT FILE6_30H.FORM_NO2 IS NULL))"
    cString = cString & turn(cString) & "FILE6_30H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_30H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_30H.CODE = " & addvalue(xCode.text)
    If ValidNum(xForm_No.text) Then
        cString = cString & turn(cString) & "FILE6_30H.FORM_NO = " & DateSq(xForm_No.text)
    End If
    
    cString = cString & " ORDER by FILE6_30H.DOC_NO"
    Set DATA12.Recordset = myRecordSet(cString, con)
    fixgrd2_install
End If
End Sub
Private Sub Myloadgrd3()
Dim cString As String
If Option1(0).Value Then
    cString = "SELECT FILE6_20H.DOC_NO,FILE6_20H.FORM_NO,FILE6_20H.CODE,FILE1_10.DESCA,FILE6_20H.TOTAL_YEAR,FILE6_20H.TOTAL_YEAR_OTHER,FILE6_20H.TOTAL_TAX,FILE6_20H.TOTAL_LATE,FILE6_20H.TOTAL,FILE6_20H.YEAR_CODE" & _
              " FROM FILE6_20H INNER JOIN FILE1_10 ON FILE6_20H.CODE = FILE1_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "FILE6_20H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_20H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_20H.CODE = " & addvalue(xCode.text)
    If ValidNum(xForm_No.text) Then
        cString = cString & turn(cString) & "FILE6_20H.FORM_NO = " & DateSq(xForm_No.text)
    End If
    cString = cString & " ORDER by FILE6_20H.DOC_NO"
    Set DATA13.Recordset = myRecordSet(cString, con)
    Fixgrd3
Else
    cString = "SELECT FILE6_30H.DOC_NO,FILE6_30H.FORM_NO,FILE6_30H.FORM_NO2,FILE6_30H.CODE,FILE2_10.DESCA,FILE6_30H.TOTAL_VALUE,FILE6_30H.CARD_VALUE,FILE6_30H.TOTAL_TAX,FILE6_30H.INTEREST,FILE6_30H.OTHER,FILE6_30H.TOTAL" & _
              " FROM FILE6_30H INNER JOIN FILE2_10 ON FILE6_30H.CODE = FILE2_10.CODE WHERE IsFawry = 0"
    cString = cString & turn(cString) & "FILE6_30H.DATE = " & DateSq(xDate.text)
    If ValidNum(xDoc_Pc.text) Then cString = cString & turn(cString) & "FILE6_30H.DOC_NO = " & addvalue(xDoc_Pc.text)
    If ValidNum(xCode.text) Then cString = cString & turn(cString) & "FILE6_30H.CODE = " & addvalue(xCode.text)
        If ValidNum(xForm_No.text) Then
        cString = cString & turn(cString) & "FILE6_30H.FORM_NO = " & DateSq(xForm_No.text)
    End If

    cString = cString & " ORDER by FILE6_30H.DOC_NO"
    Set DATA13.Recordset = myRecordSet(cString, con)
    Fixgrd3_install
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set closedayfrm = Nothing
closeCon con
End Sub
Private Sub cmdPrint_Click()
Dim cHead1 As String, cHead2 As String, cHead3 As String, nRate As Double, I As Long
cHead1 = "≈Ã„«·Ï ”œ«œ  " & IIf(Option1(0).Value, "«⁄÷«¡ ⁄«„·Ì‰", "⁄÷ÊÌ… „ﬁ”ÿ…") & " » «—ÌŒ : " & xDate.text
Dim aRow(0) As Variant
aRow(0) = AddFlag(Empty, "row", 1)
aRow(0) = AddFlag(aRow(0), "col", 0)
aRow(0) = AddFlag(aRow(0), "cols", 5)
aRow(0) = AddFlag(aRow(0), "TEXT", "«·≈Ã„«·Ì")

Dim nwidth As Double
For I = 0 To grid2.Cols - 1
    If Not grid2.ColHidden(I) Then
        nwidth = grid2.ColWidth(I) + nwidth
    End If
Next
nRate = 11500 / nwidth

PrintGrdNew.doprint grid2, nRate, -2, cHead1, cHead2, cHead3, , False, False, 10, , aRow
PrintGrdNew.Show 1
End Sub
Private Sub xExit_Click()
Unload Me
End Sub
Private Sub doprint()
End Sub
Private Sub fixgrd1()
With grid1
    .Cols = 10
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "«·√”„"
    .TextMatrix(0, 4) = "«‘ —«ﬂ «·”‰…"
    .TextMatrix(0, 5) = "«‘ —«ﬂ«  „ √Œ—…"
    .TextMatrix(0, 6) = "ﬁÌ„… „÷«›…"
    .TextMatrix(0, 7) = "€—«„…  √ŒÌ—"
    .TextMatrix(0, 8) = "≈Ã„«·Ì"
    
    .BackColorFixed = &HE0E0E0
    .ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 11
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 3 To .Cols - 2
        .Subtotal flexSTSum, -1, I, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    
    If .rows > 1 Then
        For I = 3 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If
    
End With
End Sub
Private Sub fixgrd1_install()
With grid1
    .Cols = 12
    .ColWidth(0) = 900
    .ColWidth(1) = 900
    .ColWidth(2) = 1200
    .ColWidth(3) = 1000
    .ColWidth(4) = 2500
    .ColWidth(5) = 1200
    .ColWidth(6) = 900
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
'    .ColWidth(9) = 1200
   
    '.TextMatrix(0, 5) = "«·÷—Ì»…"
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 3) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 4) = "«·√”„"
    .TextMatrix(0, 5) = "ﬁÌ„… «·ﬁ”ÿ"
    .TextMatrix(0, 6) = "ﬂ«—‰ÌÂ« "
    .TextMatrix(0, 7) = "«·÷—Ì»…"
    .TextMatrix(0, 8) = "«·›«∆œ…"
    .TextMatrix(0, 9) = " »—⁄"
    .TextMatrix(0, 10) = "„’«—Ì› «œ«—Ì…"
    .TextMatrix(0, 10 + 1) = "«·≈Ã„«·Ì"
    
    .BackColorFixed = &HC0FFC0
    
    '.ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 4 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0.00", &HE0E0E0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    
    If .rows > 1 Then
        For I = 4 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If

End With
End Sub
Private Sub fixgrd2()
With grid2
    .Cols = 10
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "«·√”„"
    .TextMatrix(0, 4) = "«‘ —«ﬂ «·”‰…"
    .TextMatrix(0, 5) = "«‘ —«ﬂ«  „ √Œ—…"
    .TextMatrix(0, 6) = "ﬁÌ„… „÷«›…"
    .TextMatrix(0, 7) = " √ŒÌ—"
    .TextMatrix(0, 8) = "≈Ã„«·Ì"
    
    .BackColorFixed = &HE0E0E0
    .ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 2
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    If .rows > 1 Then
        For I = 3 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If
End With
End Sub
Private Sub fixgrd2_install()
With grid2
    .Cols = 12
    .ColWidth(0) = 900
    .ColWidth(1) = 900
    .ColWidth(2) = 900
    .ColWidth(3) = 1200
    .ColWidth(4) = 2500
    .ColWidth(5) = 1200
    .ColWidth(6) = 900
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1200
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 3) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 4) = "«·√”„"
    .TextMatrix(0, 5) = "ﬁÌ„… «·ﬁ”ÿ"
    .TextMatrix(0, 6) = "ﬂ«—‰ÌÂ« "
    .TextMatrix(0, 7) = "«·÷—Ì»…"
    .TextMatrix(0, 8) = "«·›«∆œ…"
    .TextMatrix(0, 9) = " »—⁄"
    .TextMatrix(0, 10) = "„’«—Ì› «œ«—Ì…"
    .TextMatrix(0, 10 + 1) = "«·≈Ã„«·Ì"
    
    .BackColorFixed = &HC0FFC0
    '.ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0.00", &HE0E0E0, vbBlack, True, "«·≈Ã„«·Ì"
    Next

    If .rows > 1 Then
        For I = 5 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If

End With
End Sub
Private Sub Fixgrd3()
With grid3
    .Cols = 10
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1200
    .ColWidth(3) = 3000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1000
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "«·√”„"
    .TextMatrix(0, 4) = "«‘ —«ﬂ «·”‰…"
    .TextMatrix(0, 5) = "«‘ —«ﬂ«  „ √Œ—…"
    .TextMatrix(0, 6) = "ﬁÌ„… „÷«›…"
    .TextMatrix(0, 7) = " √ŒÌ—"
    .TextMatrix(0, 8) = "≈Ã„«·Ì"
    
    .ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 3 To .Cols - 2
        .Subtotal flexSTSum, -1, I, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    If .rows > 1 Then
        For I = 3 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If
End With
End Sub
Private Sub Fixgrd3_install()
With grid3
    .Cols = 11
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 2500
    .ColWidth(5) = 1200
    
    .ColWidth(6) = 900
    .ColWidth(7) = 1100
    .ColWidth(8) = 1050
    .ColWidth(9) = 1050
    .ColWidth(10) = 1200
    
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 2) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 3) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 4) = "«·√”„"
    .TextMatrix(0, 5) = "ﬁÌ„… «·ﬁ”ÿ"
    .TextMatrix(0, 6) = "ﬂ«—‰ÌÂ« "
    .TextMatrix(0, 7) = "«·÷—Ì»…"
    .TextMatrix(0, 8) = "«·›«∆œ…"
    .TextMatrix(0, 9) = " »—⁄"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ì"
    
    '.ColHidden(.Cols - 1) = True
    For I = 0 To .Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"
    Next
    If .rows > 1 Then
        For I = 5 To .Cols - 1
            .TextMatrix(1, I) = mRound(.TextMatrix(1, I))
        Next
    End If
End With
End Sub
Private Function myload() As Boolean
Dim cString As String
xdoc_no.text = CardTable!doc_no
xDate.text = myFormat_p(CardTable!Date)
xdate_cal.Value = myFormat_p(CardTable!Date)
xInv_count_paid.Caption = Myvalue(CardTable!INV_COUNT_PAID)
xtotal_paid.Caption = Myvalue(CardTable!Total_paid)
xtotal_net.Caption = Myvalue(mRound(CardTable!Total_paid) - mRound(CardTable!Tax_paid))
xTax.Caption = Myvalue(CardTable!Tax_paid)
xInv_count_unpaid.Caption = Myvalue(CardTable!INV_COUNT_UNPAID)
xtotal_unpaid.Caption = Myvalue(CardTable!total_Unpaid)

Handlecontrols LoadMode

myLoadGrd1
myloadgrd2
Myloadgrd3
End Function
Private Sub CmdNext_Click()
openCardTable xdoc_no.text, ">"
If CardTable.EOF Then openCardTable xdoc_no.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xdoc_no.text, "<"
If CardTable.EOF Then openCardTable xdoc_no.text
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
Private Sub Handlecontrols(nMode)
xdoc_no.Tag = nMode
aRecords = retRecords(xdoc_no.text)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")
If nMode = LoadMode Then
    xRecord_No.Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    xRecord_No.Caption = "·«  ÊÃœ ”Ã·« "
End If
cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And doc_no = ""
xdate_cal.Visible = Not IsDate(sDate)
xDate.Enabled = Not IsDate(sDate)
cmdDayBefore.Enabled = myFormat(xDate.text) = myFormat(Date)
End Sub
Private Function retRecords(pDoc_No) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pDoc_No <> "" Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN DOC_NO <= " & MyParn(pDoc_No) & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
If Option1(0).Value Then
    cString = cString & " FROM V_CLOSE_DAY " & turn(cFilter, " WHERE ") & cFilter
Else
    cString = cString & " FROM V_CLOSE_DAY_INSTALL " & turn(cFilter, " WHERE ") & cFilter
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If pDoc_No <> "" Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub xTrans_Change()
xTrans2.Caption = Format(Abs(Val(xTrans.Caption)), "fixed")
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
If Option1(0).Value Then
    cString = "SELECT TOP 1 * " & _
               " FROM  V_CLOSE_DAY"
Else
    cString = "SELECT TOP 1 * " & _
               " FROM  V_CLOSE_DAY_INSTALL"
End If

If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
End If

cFilter = ""

If IsDate(sDate) Then
    cFilter = "DATE = " & DateSq(sDate)
End If

If xDay.Value = 1 Then
    cFilter = cFilter & turn(cFilter, " AND ") & "DATE = " & DateSq(Date)
End If


If xCurrent.Value = 1 Then
    Dim aRet As Variant
    aRet = Year_Load(sSeason, , con)
    If Not IsEmpty(aRet) Then
        cFilter = cFilter & turn(cFilter, " AND ") & "DATE >= " & DateSq(retFlag(aRet, "DATE1"))
    End If
End If

If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by DOC_NO DESC"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by DOC_NO ASC"
End If


Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub myUndo(Optional bNoCellPos As Boolean = False)
If Trim(xdoc_no.text) <> "" Then
    openCardTable xdoc_no.text
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
Private Sub mydefine()
grid1.rows = 1
grid2.rows = 1
grid3.rows = 1
xdoc_no.text = ""
xInv_count_paid.Caption = ""
xInv_count_unpaid.Caption = ""
xtotal_paid.Caption = ""
xtotal_unpaid.Caption = ""
xtotal_net.Caption = ""
'xDate.text = myFormat(dSalesDate)
Handlecontrols DefineMode
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not validRow(Row) Then Exit Sub
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[form_no]", addvalue(grid1.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[USERNAME2]", addstring(cUserName & " [" & GetComputerName & "]"))
If Option1(1).Value Then
    If ValidNum(grid1.TextMatrix(Row, 2), , True) Then
        aInsert = AddFlag(aInsert, "[form_no2]", mRound(grid1.TextMatrix(Row, 2)))
    Else
        aInsert = AddFlag(aInsert, "[form_no2]", "NULL")
    End If
End If
If ValidNum(grid1.TextMatrix(Row, 1)) Then
    aInsert = AddFlag(aInsert, "[date_cashed]", "getDate()")
End If
con.BeginTrans
On Error GoTo myerror
If Option1(0).Value Then
    con.Execute addUpdate(aInsert, "FILE6_20H", "doc_no = " & addstring(grid1.TextMatrix(Row, 0)))
Else
    con.Execute addUpdate(aInsert, "FILE6_30H", "doc_no = " & addstring(grid1.TextMatrix(Row, 0)))
End If
con.CommitTrans
If Option1(0).Value Then
    Inform " „ «· ⁄œÌ· »‰Ã«Õ"
    myUndo
    If grid1.rows = 1 Then
        SendKeys "{TAB}"
    ElseIf grid1.rows = 3 Then
        CellPos 13, 1, grid1.Cols - 1
    End If
Else
    If ValidNum(grid1.TextMatrix(Row, 1), , True) And ValidNum(grid1.TextMatrix(Row, 2), , True) Then
        Inform " „ «· ⁄œÌ· »‰Ã«Õ"
        myUndo
        SendKeys "{TAB}"
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Function validRow(Row As Long) As Boolean
If Not ValidNum(grid1.TextMatrix(Row, 1)) Then
    Exit Function
Else
'    If Option1(0).Value Then
'        nFound = GetField("Select doc_no from file6_20h where form_no = " & grid1.TextMatrix(Row, 1) & " and doc_no <> " & grid1.TextMatrix(Row, 0) & " and file6_20h.date =  " & addDate(xDate.Text))
'    Else
'        If Col = 1 Then
'            nFound = GetField("Select doc_no from file6_30h where form_no = " & grid1.TextMatrix(Row, Col) & " and doc_no <> " & grid1.TextMatrix(Row, 0) & " and file6_30h.date =  " & addDate(xDate.Text))
'        ElseIf Col = 2 Then
'            nFound = GetField("Select doc_no from file6_30h where form_no2 = " & grid1.TextMatrix(Row, Col) & " and doc_no <> " & grid1.TextMatrix(Row, 0) & " and file6_30h.date =  " & addDate(xDate.Text))
'        End If
'    End If
    If Not IsEmpty(nFound) Then
        Inform "«·—ﬁ„ „ﬂ—— ›Ï „” ‰œ —ﬁ„ : " & nFound
        grid1.Select Row, Col
        Exit Function
    End If
End If
validRow = True
End Function
Private Sub grid1_EnterCell()
If Option1(0).Value Then
    If grid1.Col = 1 And grid1.Row > 1 Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
Else
    If (grid1.Col = 1 Or grid1.Col = 2) And grid1.Row > 1 Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
End If
End Sub
Private Sub Grid1_GotFocus()
grid1_EnterCell
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And ValidNum(grid1.TextMatrix(grid1.Row, 2)) And grid1.Row > 1 Then
    If MsgBox("Õ–› —ﬁ„ «·«Ì’«·", vbOKCancel + vbDefaultButton2) = vbOK Then
        con.BeginTrans
        On errror GoTo myerror
        If Option1(0).Value Then
            con.Execute "update file6_20h set file6_20h.form_no = NULL WHERE FILE6_20H.DOC_NO = " & grid1.TextMatrix(grid1.Row, 0) & " AND FILE6_20H.CLOSED = 0 and  IsFawry = 0"
        Else
            con.Execute "update file6_30h set file6_20h.form_no = NULL WHERE FILE6_20H.DOC_NO = " & grid1.TextMatrix(grid1.Row, 0) & " AND FILE6_20H.CLOSED = 0"
        End If
        con.CommitTrans
        myLoadGrd1
        Inform " „ Õ–› «·—ﬁ„ »‰Ã«Õ"
    End If
End If
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
If Col = 1 Then
    If Not ValidNum(grid1.EditText) Then
        Cancel = True
    ElseIf ValidNum(grid1.EditText) Then
'        If Option1(0).Value Then
'            nFound = GetField("Select doc_no from file6_20h where form_no = " & grid1.EditText & " and doc_no <> " & grid1.TextMatrix(Row, 0) & " and file6_20h.year_code =  " & grid1.TextMatrix(Row, grid1.Cols - 1))
'        Else
'            nFound = GetField("Select doc_no from file6_30h where form_no = " & grid1.EditText & " and doc_no <> " & grid1.TextMatrix(Row, 0) & " and year(file6_30h.date) = " & Year(xDate.Text))
'        End If
        If Not IsEmpty(nFound) Then
            Inform "«·—ﬁ„ „ﬂ—— ›Ï „” ‰œ —ﬁ„ : " & nFound
            Cancel = True
        End If
    End If
ElseIf Col = 2 Then
    If Option1(0).Value Then
        If Not ValidNum(grid1.EditText, , True) Then
            MsgBox ("«·—ﬁ„ €Ì— ’ÕÌÕ")
            Cancel = True
        End If
    Else
        If Not ValidNum(grid1.EditText) Then
            MsgBox ("«·—ﬁ„ €Ì— ’ÕÌÕ")
            Cancel = True
        End If
    End If
End If
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid2.Row > 1 Then
    If MsgBox("Õ–› —ﬁ„ «·«Ì’«·", vbOKCancel + vbDefaultButton2) = vbOK Then
        con.BeginTrans
        On errror GoTo myerror
        If Option1(0).Value Then
            con.Execute "update file6_20h " & _
                        " SET file6_20h.form_no = NULL " & _
                        ",[USERNAME2] =  " & addstring(cUserName & " [" & GetComputerName & "]") & _
                        " WHERE FILE6_20H.DOC_NO = " & grid2.TextMatrix(grid2.Row, 0) & _
                        " AND FILE6_20H.CLOSED = 0 AND IsFawry = 0"
        Else
            con.Execute "update file6_30h " & _
                        " set file6_30h.form_no = NULL " & _
                        ", FILE6_30H.FORM_NO2 = NULL " & _
                        ",[USERNAME2] =  " & addstring(cUserName & " [" & GetComputerName & "]") & _
                        " WHERE FILE6_30H.DOC_NO = " & grid2.TextMatrix(grid2.Row, 0) & _
                        " AND FILE6_30H.CLOSED = 0"
        End If
        con.CommitTrans
        'myloadgrd2
        Inform " „ Õ–› «·—ﬁ„ »‰Ã«Õ"
        myUndo
        grid2.SetFocus
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub Option1_Click(Index As Integer)
myUndo
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If Option1(0).Value Then
        MemberLookupAll Me, oSearch
    Else
        Member_InLookupAll Me, oSearch
    End If
ElseIf KeyCode = 13 Then
    cmdGo_Click
    If ValidNum(xCode.text) And grid1.rows > 1 Then
        grid1.SetFocus
    End If
End If
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If ValidNum(xCode.text) Then
    Dim aMember As Variant
    If Option1(0).Value Then
        aMember = Member_Load(xCode.text, , con)
        If Not IsEmpty(aMember) Then
            xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
        Else
            xCodeDesca.Caption = ""
        End If
    Else
        aMember = Member_Load_install(xCode.text, , con)
        'aPaid = Member_Paid_Install(xCode.Text, , con)
        If Not IsEmpty(aMember) Then
            xCodeDesca.Caption = retFlag(aMember, "Desca") & ""
        Else
            xCodeDesca.Caption = ""
        End If
    End If
Else
    xCodeDesca.Caption = ""
End If
End Sub

Private Sub xCurrent_Click()
'openCardTable
'Handlecontrols LoadMode
myUndo
End Sub

Private Sub xdate_cal_Click()
Dim aRet As Variant, cString As String
If Option1(0).Value Then
    cString = "Select top 1 doc_no from V_CLOSE_DAY where date = " & DateSq(xdate_cal.Value)
Else
    cString = "Select top 1 doc_no from V_CLOSE_DAY_INSTALL where date = " & DateSq(xdate_cal.Value)
End If
If cFilter <> "" Then
    cString = cString & " and " & cFilter
End If
aRet = GetField(cString, con)
If Not IsEmpty(aRet) Then
    xdoc_no.text = aRet
    myUndo
Else
    xDate.text = myFormat_p(xdate_cal.Value)
    mydefine
End If
End Sub
Private Sub xDate_DblClick()
Set datefrm.oDate = xDate
datefrm.Show 1
End Sub
Private Sub xdate_LostFocus()
Dim aRet As Variant, cString As String
xDate.text = myFormat_p(xDate.text)
If Not IsDate(xDate.text) Then
    myUndo
Else
    If Option1(0).Value Then
        cString = "Select top 1 doc_no from V_CLOSE_DAY where date = " & DateSq(xDate.text)
    Else
        cString = "Select top 1 doc_no from V_CLOSE_DAY_INSTALL where date = " & DateSq(xDate.text)
    End If
    If cFilter <> "" Then cString = cString & " and " & cFilter
    aRet = GetField(cString, con)
    If Not IsEmpty(aRet) Then
        xdoc_no.text = aRet
        myUndo
    Else
        xdate_cal.Value = myFormat_p(xDate.text)
        mydefine
    End If
End If
End Sub
Private Sub xDay_Click()
myUndo
End Sub
Sub myProc()
    If ActiveControl.Name = xCode.Name Then
        xCode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
        xCodeDesca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
        oSearch.Hide
        cmdGo_Click
    ElseIf ActiveControl.Name = cmdDayBefore.Name Then
        'If Option1(0).Value Then
            If restoreDoc(oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0), IIf(Option1(0).Value, "FILE6_20H", "FILE6_30H")) Then
                Inform " „ «· ÕÊÌ· »‰Ã«Õ"
                myUndo
                'Unload oSearchDoc
            End If
        'End If
    End If
End Sub
Private Function restoreDoc(pDoc_No As String, pFile As String) As Boolean
con.BeginTrans
On Error GoTo myerror
con.Execute "UPDATE " & pFile & " set [date] = " & addDate(xDate.text) & " WHERE DOC_NO = " & addvalue(pDoc_No)
con.CommitTrans
restoreDoc = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Sub xDoc_Pc_GotFocus()
myGotFocus xDoc_Pc
End Sub
Private Sub xDoc_Pc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdGo_Click
    If ValidNum(xDoc_Pc.text) And grid1.rows > 1 Then
        grid1.SetFocus
    ElseIf grid2.rows > 1 Then
        grid2.SetFocus
    End If
End If
End Sub
Private Sub xDoc_Pc_LostFocus()
myLostFocus xDoc_Pc
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
With grid1
If Option1(0).Value Then
    If Col < .Cols - 9 Then
        .Col = Col + 1
    ElseIf Row < .rows - 1 Then
        .ShowCell .Row, 0
        .Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
    Else
        .Select Row, Col
    End If
Else
    If Col < .Cols - 7 Then
        .Col = Col + 1
    ElseIf Row < .rows - 1 Then
        .ShowCell .Row, 0
        .Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
    Else
        .Select Row, Col
    End If
End If
End With
End Sub

Private Sub xForm_No_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    myloadgrd2
    Myloadgrd3
    If ValidNum(xForm_No.text) And grid2.rows > 1 Then
        grid2.SetFocus
    End If
End If
End Sub

