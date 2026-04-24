VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form carsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·”Ì«—« "
   ClientHeight    =   8715
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   13695
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   3840
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   1755
      Width           =   7260
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3420
         Width           =   330
      End
      Begin VB.TextBox xGas3 
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
         Height          =   375
         Left            =   90
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Tag             =   "MM"
         Top             =   2655
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox xGas2 
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
         Height          =   375
         Left            =   90
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Tag             =   "MM"
         Top             =   2250
         Width           =   1725
      End
      Begin VB.TextBox xGas1 
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
         Height          =   375
         Left            =   3555
         MaxLength       =   18
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Tag             =   "ss"
         Top             =   2655
         Width           =   1725
      End
      Begin VB.TextBox xDistance 
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
         Height          =   375
         Left            =   3555
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Tag             =   "ss"
         Top             =   2250
         Width           =   1725
      End
      Begin VB.TextBox xLine 
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
         Height          =   1725
         Left            =   90
         MaxLength       =   500
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   450
         Width           =   7080
      End
      Begin MSDataListLib.DataCombo xType_Gas 
         Height          =   315
         Left            =   2340
         TabIndex        =   17
         Top             =   3420
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÊﬁÊœ"
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
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   3465
         Width           =   825
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "“Ì  ›·›Ì‰ ·ﬂ· ﬂÌ·Ê"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   2745
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "“Ì  „Ê Ê— ·ﬂ· ﬂÌ·Ê"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   2295
         Width           =   1470
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì «” Â·«ﬂ ÊﬁÊœ"
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
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   3105
         Width           =   1635
      End
      Begin VB.Label xQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   345
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   3060
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "· — »‰“Ì‰ ·ﬂ· ﬂÌ·Ê „ —"
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
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2700
         Width           =   1620
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”«›… «·Œÿ »«·ﬂÌ·Ê „ —"
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
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   2295
         Width           =   1710
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Œÿ ”Ì— «·”Ì«—…"
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   135
         Width           =   1245
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   7785
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   50
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
         Picture         =   "Carsfrm.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Carsfrm.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   51
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
         Picture         =   "Carsfrm.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Carsfrm.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   52
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
         Picture         =   "Carsfrm.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Carsfrm.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   53
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
         Picture         =   "Carsfrm.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Carsfrm.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   6390
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "Carsfrm.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Carsfrm.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Carsfrm.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Carsfrm.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Carsfrm.frx":1A6FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
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
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Carsfrm.frx":1CC73
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  ≈÷«›Ì…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   5625
      Width           =   13515
      Begin VB.TextBox xPrice 
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
         Left            =   9810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Tag             =   "D"
         Top             =   210
         Width           =   2130
      End
      Begin VB.CheckBox xStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„ Êﬁ›…"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1080
         Width           =   870
      End
      Begin VB.TextBox xRemark 
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
         Height          =   690
         Left            =   -45
         MaxLength       =   1000
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1395
         Width           =   11985
      End
      Begin VB.TextBox xColor 
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
         Height          =   360
         Left            =   9000
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   2940
      End
      Begin VB.TextBox xDate_tax 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Tag             =   "D"
         Top             =   585
         Width           =   2130
      End
      Begin VB.TextBox xDate_End 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Tag             =   "D"
         Top             =   990
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
         Height          =   345
         Left            =   9810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "D"
         Top             =   990
         Width           =   2130
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·”Ì«—…"
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
         Left            =   12015
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Â«Ì… «·÷—Ì»…"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   12045
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1350
         Width           =   660
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»œ«Ì… ⁄„· «·”Ì«—…"
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
         Left            =   12045
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Â«Ì… ⁄„· «·”Ì«—…"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«··Ê‰"
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
         Left            =   12045
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4920
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   675
      Width           =   6180
      Begin VB.TextBox xdesca 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   3930
      End
      Begin VB.TextBox xYear_make 
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
         Height          =   375
         Left            =   2160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   4095
         Width           =   1905
      End
      Begin VB.TextBox xModel 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   3930
      End
      Begin VB.TextBox xTraffic 
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
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3930
      End
      Begin VB.TextBox xDate_Auth_End 
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
         Left            =   2160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "D"
         Top             =   3735
         Width           =   1905
      End
      Begin VB.TextBox xDate_Auth 
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
         Height          =   375
         Left            =   2160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "D"
         Top             =   3330
         Width           =   1905
      End
      Begin VB.TextBox xGovern 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2925
         Width           =   3930
      End
      Begin VB.TextBox xMotor 
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
         Height          =   375
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   3930
      End
      Begin VB.TextBox xBoard 
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
         Height          =   360
         Left            =   135
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1395
         Width           =   3930
      End
      Begin VB.TextBox xBody 
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
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2565
         Width           =   3930
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   2745
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "N"
         Top             =   180
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xDriver 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Tag             =   "ss"
         Top             =   4500
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "«·”«∆ﬁ"
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
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
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
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "”‰… «·’‰⁄"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   4185
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰Ê⁄ Ê«·„ÊœÌ·"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„—Ê—"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1830
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «‰ Â«¡ «· —ŒÌ’"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3780
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ŒÌ’"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3375
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Õ«›Ÿ…"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3015
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„Ê Ê—"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2205
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «··ÊÕ…"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1485
         Width           =   795
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·‘«”Ì…"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2610
         Width           =   900
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1530
      Top             =   225
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1800
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   3555
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
   Begin MSAdodcLib.Adodc DATA12 
      Height          =   330
      Left            =   3555
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
   Begin MSAdodcLib.Adodc DATA13 
      Height          =   330
      Left            =   3555
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
Attribute VB_Name = "CarsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public bEdit As Boolean
Dim oSearch As New Search3
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdGroup_Click()
Dim oFlag As New flag_mainfrm, sCode As String
sCode = xType_Gas.BoundText
oFlag.sTable = "TYPE_GAS_CODES"
oFlag.sCaption = "«ﬁ”«„ «·«’‰«›"
oFlag.nZero = -1
oFlag.bEdit = True
oFlag.Show 1
DATA2.Refresh
xType_Gas.BoundText = sCode
xType_Gas_LostFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then KeyAscii = 0
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    CmdSave_Click
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End If
End Sub

Private Sub Form_Load()
openCon con

Set DATA1.Recordset = myRecordSet("SELECT * FROM DRIVER WHERE DRIVER = 1 ORDER BY DESCA", con)
Set xDriver.RowSource = DATA1
xDriver.ListField = "Desca"
xDriver.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "TYPE_GAS_CODES"
Set xType_Gas.RowSource = DATA2
xType_Gas.ListField = "Desca"
xType_Gas.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
xDesca.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    con.BeginTrans
    con.Execute "Delete  From FILE0_50  Where code = " & MyParn(xCode.Text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
    Else
        mydefine
    End If
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
carsLookupAll Me, oSearch
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
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
cmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Text = Newflag("CARS", "CODE")
xBoard.Text = ""
xTraffic.Text = ""
xBody.Text = ""
xMotor.Text = ""
xGovern.Text = ""
xModel.Text = ""
xYear_make.Text = ""
xDesca.Text = ""
xLine.Text = ""
xDate_Auth.Text = ""
xDate_Auth_End.Text = ""
xType_Gas.BoundText = ""
xStop.Value = 0
xDate_tax.Text = ""
xDate_Begin.Text = ""
xColor.Text = ""
xDate_End.Text = ""
xRemark.Text = ""
xGas1.Text = ""
xDriver.BoundText = ""
xLine.Text = ""
xQuant.Caption = ""
xDistance.Text = ""
xGas2.Text = ""
xGas3.Text = ""
xPrice.Text = ""
Handlecontrols DefineMode
End Sub
Private Sub MyLoad()
xCode.Text = CardTable!CODE & ""
xBoard.Text = CardTable!Board & ""
xTraffic.Text = CardTable!Traffic & ""
xDriver.BoundText = CardTable!Driver & ""
xDistance.Text = Myvalue(CardTable!Distance)
xLine.Text = CardTable!Line & ""
xBody.Text = CardTable!Body & ""
xMotor.Text = CardTable!Motor & ""
xGovern.Text = CardTable!Govern & ""
xModel.Text = CardTable!Model & ""
xYear_make.Text = CardTable!Year_make & ""
xDesca.Text = CardTable!DESCA & ""
xDate_Auth.Text = myFormat(CardTable!Date_Auth)
xDate_Auth_End.Text = myFormat(CardTable!Date_Auth_End)
xType_Gas.BoundText = CardTable!Type_Gas & ""
xStop.Value = IIf(CardTable!Stop, 1, 0)
xDate_tax.Text = CardTable!Date_tax & ""
xDate_Begin.Text = myFormat(CardTable!DATE_BEGIN)
xColor.Text = CardTable!Color & ""
xDate_End.Text = myFormat(CardTable!DATE_END)
xRemark.Text = CardTable!Remark & ""
xGas1.Text = Myvalue(CardTable!gas1)
xGas2.Text = Myvalue(CardTable!gas2)
xGas3.Text = Myvalue(CardTable!gas3)
xPrice.Text = Myvalue(CardTable!PRICE)
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
Calctotals
End Sub
Private Function MyReplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "Board", addstring(xBoard.Text))
aInsert = AddFlag(aInsert, "Traffic", addstring(xTraffic.Text))
aInsert = AddFlag(aInsert, "Body", addstring(xBody.Text))
aInsert = AddFlag(aInsert, "Driver", addstring(xDriver.BoundText))
aInsert = AddFlag(aInsert, "Line", addstring(xLine.Text))
aInsert = AddFlag(aInsert, "Distance", Val(xDistance.Text))
aInsert = AddFlag(aInsert, "Motor", addstring(xMotor.Text))
aInsert = AddFlag(aInsert, "Govern", addstring(xGovern.Text))
aInsert = AddFlag(aInsert, "Model", addstring(xModel.Text))
aInsert = AddFlag(aInsert, "Year_make", addvalue(xYear_make.Text))
aInsert = AddFlag(aInsert, "Desca", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "Date_Auth", addDate(xDate_Auth.Text))
aInsert = AddFlag(aInsert, "Date_Auth_End", addDate(xDate_Auth_End.Text))
aInsert = AddFlag(aInsert, "Type_Gas", addvalue(xType_Gas.BoundText))
aInsert = AddFlag(aInsert, "Stop", xStop.Value)
aInsert = AddFlag(aInsert, "Date_tax", addDate(xDate_tax.Text))
aInsert = AddFlag(aInsert, "Date_Begin", addDate(xDate_Begin.Text))
aInsert = AddFlag(aInsert, "Color", addstring(xColor.Text))
aInsert = AddFlag(aInsert, "Date_End", addDate(xDate_End.Text))
aInsert = AddFlag(aInsert, "Remark", addstring(xRemark.Text))
aInsert = AddFlag(aInsert, "gas1", Val(xGas1.Text))
aInsert = AddFlag(aInsert, "gas2", Val(xGas2.Text))
aInsert = AddFlag(aInsert, "gas3", Val(xGas3.Text))
aInsert = AddFlag(aInsert, "[PRICE]", Val(xPrice.Text))
con.BeginTrans
On Error GoTo myError
If xCode.Tag = DefineMode Then
   xCode.Text = Newflag("CARS", "CODE", con)
   aInsert = AddFlag(aInsert, "CODE", addvalue(xCode.Text))
   con.Execute addInsert(aInsert, "CARS")
Else
    con.Execute addUpdate(aInsert, "CARS", "CODE = " & addvalue(xCode.Text))
End If
con.CommitTrans
MyReplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()

xCode.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
If MsgBox(cMsgExit, vbOKCancel + vbDefaultButton1) = vbOK Then
    CmdSave_Click
End If
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload oSearch
Set oSearch = Nothing
Err.Clear
closeCon con
End Sub

Private Sub xCode_LostFocus()
myLostFocus xCode
If Not IsNumeric(xCode.Text) Then Exit Sub
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If (Not CardTable.EOF) Then
    MyLoad
ElseIf xCode.Tag = LoadMode Then
    mydefine
End If
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "«·ﬂÊœ €Ì— „”Ã·"
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
    MyLoad
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT CARS.* FROM CARS"
cString = cString & " ORDER BY CARS.[CODE]"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Private Sub xColor_GotFocus()
myGotFocus xColor
End Sub
Private Sub xColor_LostFocus()
myLostFocus xColor
End Sub
Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
Private Sub xDate_tax_GotFocus()
myGotFocus xDate_tax
End Sub
Private Sub xDate_tax_LostFocus()
myLostFocus xDate_tax
myValidDate xDate_tax
End Sub
Private Sub xDate_End_GotFocus()
myGotFocus xDate_End
End Sub
Private Sub xDate_End_LostFocus()
myLostFocus xDate_End
myValidDate xDate_End
End Sub
Private Sub xDate_Begin_GotFocus()
myGotFocus xDate_Begin
End Sub
Private Sub xDate_Begin_LostFocus()
myLostFocus xDate_Begin
End Sub
Private Sub xRemark_GotFocus()
myGotFocus xRemark
End Sub
Private Sub xRemark_LostFocus()
myLostFocus xRemark
End Sub
Private Sub xType_Gas_GotFocus()
If Not xType_Gas.MatchedWithList Then xType_Gas.BoundText = ""
myGotFocus xType_Gas
End Sub
Private Sub xType_Gas_LostFocus()
myLostFocus xType_Gas
End Sub
Private Sub xOwner_GotFocus()
myGotFocus xOwner
End Sub
Private Sub xOwner_LostFocus()
myLostFocus xOwner
End Sub
Private Sub xWeigth_GotFocus()
myGotFocus xWeigth
End Sub
Private Sub xWeigth_LostFocus()
myLostFocus xWeigth
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
End Sub
Private Sub xYear_make_GotFocus()
myGotFocus xYear_make
End Sub
Private Sub xYear_make_LostFocus()
myLostFocus xYear_make
End Sub
Private Sub xModel_GotFocus()
myGotFocus xModel
End Sub
Private Sub xModel_LostFocus()
myLostFocus xModel
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xTraffic_GotFocus()
myGotFocus xTraffic
End Sub
Private Sub xTraffic_LostFocus()
myLostFocus xTraffic
End Sub
Private Sub xDate_Auth_End_GotFocus()
myGotFocus xDate_Auth_End
End Sub
Private Sub xDate_Auth_End_LostFocus()
myLostFocus xDate_Auth_End
myValidDate xDate_Auth_End
End Sub
Private Sub xDate_Auth_GotFocus()
myGotFocus xDate_Auth
End Sub
Private Sub xDate_Auth_LostFocus()
myLostFocus xDate_Auth
myValidDate xDate_Auth
End Sub
Private Sub xGovern_GotFocus()
myGotFocus xGovern
End Sub
Private Sub xGovern_LostFocus()
myLostFocus xGovern
End Sub
Private Sub xMotor_GotFocus()
myGotFocus xMotor
End Sub
Private Sub xMotor_LostFocus()
myLostFocus xMotor
End Sub
Private Sub xBoard_GotFocus()
myGotFocus xBoard
End Sub
Private Sub xBoard_LostFocus()
myLostFocus xBoard
End Sub
Private Sub xBody_GotFocus()
myGotFocus xBody
End Sub
Private Sub xBody_LostFocus()
myLostFocus xBody
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xgas1_GotFocus()
myGotFocus xGas1
End Sub
Private Sub xgas1_LostFocus()
myLostFocus xGas1
Calctotals
End Sub
Private Sub xDistance_GotFocus()
myGotFocus xDistance
End Sub
Private Sub xDistance_LostFocus()
myLostFocus xDistance
Calctotals
End Sub
Private Sub xLine_GotFocus()
myGotFocus xLine
End Sub
Private Sub xLine_LostFocus()
myLostFocus xLine
End Sub
Private Sub xDriver_GotFocus()
myGotFocus xDriver
End Sub
Private Sub xDriver_LostFocus()
myLostFocus xDriver
If Not xDriver.MatchedWithList Then xDriver.BoundText = ""
End Sub
Private Sub Calctotals()
xQuant.Caption = Myvalue(Round(Val(xDistance.Text) * Val(Me.xGas1.Text), 2))
End Sub
Private Sub xGas3_GotFocus()
myGotFocus xGas3
End Sub
Private Sub xGas3_LostFocus()
myLostFocus xGas3
End Sub
Private Sub xGas2_GotFocus()
myGotFocus xGas2
End Sub
Private Sub xGas2_LostFocus()
myLostFocus xGas2
End Sub
