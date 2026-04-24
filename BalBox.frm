VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BalBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   12975
   Begin VB.Frame Frame2 
      Caption         =   "’«œ—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   1485
      Width           =   3480
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "⁄Ã“ :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2835
         Width           =   450
      End
      Begin VB.Label xShort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2790
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì ’«œ— :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3375
         Width           =   1140
      End
      Begin VB.Label xOut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   3330
         Width           =   1365
      End
      Begin VB.Line Line2 
         X1              =   45
         X2              =   3420
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   45
         X2              =   3420
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label xChq_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2430
         Width           =   1365
      End
      Begin VB.Label xCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label xPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label xPurchase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label xCash_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label xBank_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label xTrans_out 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄«  ‰ﬁœÌ… :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‘ —Ì«  ‰ﬁœÌ… :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„’«—Ì› :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2115
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ìœ«⁄«  »‰ﬂÌ… :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì  ÕÊÌ·«  „‰ :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1800
         Width           =   1590
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "√Ê—«ﬁ œ›⁄ :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2475
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì Ã«—Ï «·‘—ﬂ«¡ :"
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ê«—œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   9135
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2295
      Width           =   3750
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„»Ì⁄«  ‰ﬁœÌ… :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì—«œ«  :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1710
         Width           =   675
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ÕÊÌ· ≈·Ì :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁ»Ê÷«  ‰ﬁœÌ… :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ê—«ﬁ ﬁ»÷ :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2025
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”ÕÊ»«  »‰ﬂÌ… :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   990
         Width           =   1245
      End
      Begin VB.Line Line3 
         X1              =   225
         X2              =   3600
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   180
         X2              =   3555
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì Ê—«œ :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Label xIn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2565
         Width           =   1365
      End
      Begin VB.Label xChq_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2025
         Width           =   1365
      End
      Begin VB.Label xIncome 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1665
         Width           =   1365
      End
      Begin VB.Label xTrans_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   1365
      End
      Begin VB.Label xBank_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label xCash_in 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label xSales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   270
      Width           =   3705
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   345
         Left            =   135
         TabIndex        =   16
         Top             =   900
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“‰… :"
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
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ :"
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
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Height          =   270
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5850
      Width           =   7260
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "BalBox.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   4410
         Picture         =   "BalBox.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1410
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   " ›’Ì·Ì «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton Cmd_Print2 
         Caption         =   "≈Ã„«·Ì Õ—ﬂ… ÌÊ„Ì…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1545
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   5820
         Picture         =   "BalBox.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1410
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1620
      Width           =   3705
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ ”«»ﬁ :"
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label xFirst_Balance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   9135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5265
      Width           =   3750
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «Œ— «·ÌÊ„ :"
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label xLast_Balance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1365
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3750
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -585
      Top             =   1710
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   5550
      _cx             =   9790
      _cy             =   11456
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
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
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
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
End
Attribute VB_Name = "BalBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
If Not IsDate(xdate1.Text) Then
    MsgBox " «—ÌŒ «·«Ê· ÷—Ê—Ì"
    Exit Sub
End If
If XBOX.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

'--------------  Ê«—œ
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as First_Balance"

cField2 = myiif( _
        "( FLAG = 0)", "PLUS - MINUS") & _
        " As First_Bal"
                                  
cField3 = myiif( _
        " (FLAG = 9 or Flag = 10 )", "PLUS - MINUS") & _
        " As Sales"

cField4 = myiif( _
        " (FLAG = 1 or Flag = 4 )", "PLUS - MINUS") & _
        " As Cash_In"

cField5 = myiif( _
        " (FLAG = 14)", "PLUS") & _
        " As Bank_In"

cField6 = myiif( _
        " (FLAG = 8)", "PLUS") & _
        " As Trans_In"

cField7 = myiif( _
        " (FLAG = 6)", "PLUS ") & _
        " As Income"

cField8 = myiif( _
        " (FLAG = 13)", "PLUS") & _
        " As Chq_In"

' ----------- ’«œ—
cField9 = myiif( _
        " (FLAG = 11 or flag = 12)", "MINUS - PLUS") & _
        " As Purchase"

cField10 = myiif( _
        " (FLAG = 2 or flag = 3)", "MINUS - PLUS") & _
        " As Cash_out"

cField11 = myiif( _
        " (FLAG = 15)", "MINUS") & _
        " As BANK_OUT"

cField12 = myiif( _
        " (FLAG = 7)", "MINUS") & _
        " As TRANS_OUT"

cField13 = myiif( _
        " (FLAG = 5)", "MINUS") & _
        " As CHARGES"
        
cField14 = myiif( _
        " (FLAG = 16)", "MINUS") & _
        " As CHQ_OUT"
        
cField15 = myiif( _
        " (FLAG = 17)", "MINUS") & _
        " As PART"
        
cField16 = myiif( _
        " (FLAG = 21)", "MINUS") & _
        " As SHORT"
        
        
cwhere = ""
If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

cField17 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as Last_Balance"

cString = "Select " & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & _
           cField6 & "," & cField7 & "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & "," & cField15 & "," & cField16 & "," & cField17 & _
           " From BOXMOVE "

cwhere = ""
If IsDate(xdate1.Text) Then cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)

If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " BOX = " & MyParn(XBOX.BoundText)
End If

cString = cString & turn(cwhere) & cwhere
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (sourcetable.EOF And sourcetable.BOF) Then
    xFirst_balance.Caption = Format(Val(sourcetable!First_Balance & "") + Val(sourcetable!First_Bal & ""), "FIXED")
    xSales.Caption = Format(sourcetable!sales, "Fixed")
    xCash_in.Caption = Format(sourcetable!Cash_in, "Fixed")
    xChq_in.Caption = Format(sourcetable!Chq_in, "Fixed")
    xBank_in.Caption = Format(sourcetable!Bank_in, "Fixed")
    xTrans_in.Caption = Format(sourcetable!Trans_In, "Fixed")
    xIncome.Caption = Format(sourcetable!Income, "Fixed")
    
    xPurchase.Caption = Format(sourcetable!purchase, "Fixed")
    xCash_out.Caption = Format(sourcetable!Cash_out, "Fixed")
    xBank_out.Caption = Format(sourcetable!Bank_Out, "Fixed")
    xTrans_out.Caption = Format(sourcetable!Trans_Out, "Fixed")
    xCharges.Caption = Format(sourcetable!CHARGES, "Fixed")
    xChq_out.Caption = Format(sourcetable!Chq_Out, "Fixed")
    xPart.Caption = Format(sourcetable!Part, "Fixed")
    xShort.Caption = Format(sourcetable!Short, "Fixed")
    
    xIn.Caption = Format(Val(xFirst_balance.Caption) + Val(xSales.Caption) + Val(xCash_in.Caption) + Val(xBank_in.Caption) + Val(xTrans_in.Caption) + Val(xIncome.Caption), "FIXED")
    xOut.Caption = Format(Val(xPurchase.Caption) + Val(xCash_out.Caption) + Val(xTrans_out.Caption) + Val(xCharges.Caption) + Val(xBank_out.Caption) + Val(xChq_out.Caption) + Val(xPart.Caption), "FIXED") + Round(Val(xShort.Caption), 2)
    xLast_Balance.Caption = Format(Val(sourcetable!Last_Balance & ""), "FIXED")
End If
sourcetable.Close
Set sourcetable = Nothing
myloadgrd
End Sub
Private Sub cmdPrint_Click()
Dim temptable As New ADODB.Recordset, aHeader(1)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
contemp.BeginTrans
If IsDate(xdate1.Text) Or IsDate(XDATE2.Text) Then
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If
If XBOX.MatchedWithList Then
    aHeader(0) = "≈Ã„«·Ì —’Ìœ Œ“‰… : " & XBOX.Text
End If

temptable.AddNew
temptable!str21 = TurnValue(aHeader(0))

temptable!Date1 = DateFix(xdate1.Text)
temptable!date2 = DateFix(XDATE2.Text)


temptable!val1 = Val(xFirst_balance.Caption)
temptable!val2 = Val(xSales.Caption)
temptable!val3 = Val(xCash_in.Caption)
temptable!val4 = Val(xBank_in.Caption)
temptable!val5 = Val(xTrans_in.Caption)
temptable!Val6 = Val(xIncome.Caption)
temptable!Val7 = Val(xChq_in.Caption)

temptable!Val8 = Val(xPurchase.Caption)
temptable!val9 = Val(xCash_out.Caption)
temptable!val10 = Val(xBank_out.Caption)
temptable!val11 = Val(xTrans_out.Caption)
temptable!val12 = Val(xCharges.Caption)
temptable!val13 = Val(xShort.Caption)
temptable!Val16 = Val(xIn.Caption)
temptable!Val17 = Val(xOut.Caption)
temptable!Val18 = Val(xLast_Balance.Caption)
temptable!Str11 = TurnValue(retHeader(aHeader, 0, 1))
temptable!str12 = TurnValue(retHeader(aHeader, 1, 1))
temptable.Update
contemp.CommitTrans

main.Report1.ReportFileName = App.Path & "\Reports\BALBOX.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1

temptable.Close
Set temptable = Nothing
End Sub
Private Sub Cmd_Print2_Click()
doprint2

End Sub
Private Sub CmdPrint2_Click()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
Dim temptable As New ADODB.Recordset
Dim aHeader(2)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If XBOX.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as FirstBalance"

cwhere = ""
If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

cString = "Select  BOXMOVE.*," & cField1 & _
           " From boxmove"

cwhere = ""
If IsDate(xdate1.Text) Then
    cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If

If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE >= " & DateSq(xdate1.Text)
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
    aHeader(0) = "—’Ìœ «·Œ“‰… : " & XBOX.Text
End If

cString = cString & turn(cwhere) & cwhere & "  ORDER BY DATE,Flag"
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.EOF) Then
    nBalance = Val(sourcetable!FirstBalance & "")
    If nBalance <> 0 Then
        temptable.AddNew
        temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & XBOX.Text
        temptable!str2 = "—’Ìœ ”«»ﬁ"
        temptable!val1 = nBalance
        temptable!val2 = 0
        temptable!val3 = nBalance
        temptable!str21 = retHeader(aHeader, 0, 1)
        temptable.Update
    End If
End If

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!Date1 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !Desca
    temptable!str3 = !CodeDesca
    temptable!val1 = !PLUS
    temptable!val2 = !MINUS
    nBalance = nBalance + Val(!PLUS & "") - Val(!MINUS & "")
    temptable!val3 = nBalance
    temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & XBOX.Text
    temptable!str21 = retHeader(aHeader, 0, 1)
    temptable.Update
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\Box2.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
sourcetable.Close: Set sourcetable = Nothing
temptable.Close: Set temptable = Nothing
End Sub
Private Sub Form_Load()
openCon con
Set boxtable = New ADODB.Recordset
boxtable.Open "FILE0_50", con, adOpenKeyset, adLockOptimistic, adCmdTable

data1.ConnectionString = strCon
data1.RecordSource = "FILE0_50"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"
xdate1.Text = RetSetting("date_first", TempSave(Me))
XDATE2.Text = Format(Date, "YYYY-MM-DD")
Fixgrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
addSetting "date_first", Format(xdate1.Text, "YYYY-MM-DD"), TempSave(Me)
closeCon con
End Sub

Private Sub GRID1_Click()
'If grid1.row <> 0 Then balBoxDtlfrm.Show 1
End Sub
Private Function myValid3() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If Not IsDate(XDATE2.Text) And Trim(XDATE2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
If Trim(XBOX.BoundText) = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Function
End If
myValid3 = True
End Function
Private Sub myloadgrd()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
If XBOX.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " BOX = " & MyParn(XBOX.BoundText)
End If

cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as FirstBalance"


If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

cString = "Select Date, " & cField1 & ", Sum(PLUS) as sumofPlus,sum(Minus) as sumofMinus,Sum(Plus- MINUS) as SumOfValue " & _
           " From boxmove "

cwhere = ""
If IsDate(xdate1.Text) Then cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)
If IsDate(XDATE2.Text) Then cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
If XBOX.BoundText <> "" Then cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
cString = cString & turn(cwhere) & cwhere & " Group by Date ORDER BY DATE"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

grid1.Rows = 1
If Not (sourcetable.EOF And sourcetable.EOF) Then
    nBalance = Val(sourcetable!FirstBalance & "")
    If Val(nBalance) <> 0 Then
        grid1.AddItem ""
        grid1.TextMatrix(grid1.Rows - 1, 0) = "—’Ìœ ”«»ﬁ"
        grid1.TextMatrix(grid1.Rows - 1, 2) = Format(Val(nBalance), "Fixed")
    End If
End If

Do Until sourcetable.EOF
    If Val(sourcetable!Sumofvalue & "") <> 0 Then
        grid1.AddItem ""
        grid1.TextMatrix(grid1.Rows - 1, 0) = sourcetable!Date
        grid1.TextMatrix(grid1.Rows - 1, 1) = Format(sourcetable!SumofPlus, "fixed")
        grid1.TextMatrix(grid1.Rows - 1, 2) = Format(sourcetable!SumofMinus, "fixed")
        grid1.TextMatrix(grid1.Rows - 1, 3) = Format(sourcetable!Sumofvalue, "fixed")
        nBalance = Val(sourcetable!Sumofvalue & "") + nBalance
        grid1.TextMatrix(grid1.Rows - 1, 4) = Format(nBalance, "fixed")
    End If
    sourcetable.MoveNext
Loop

sourcetable.Close
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
Dim temptable As New ADODB.Recordset
Dim aHeader(2)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If XBOX.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xdate1.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE < " & DateSq(xdate1.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
End If

'--------------  Ê«—œ
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cwhere) & cwhere & _
           ") as First_Balance"
cField2 = myiif( _
        "( FLAG = 0)", "PLUS - MINUS") & _
        " As First_Bal"
                                  
cField3 = myiif( _
        " (FLAG = 9 or Flag = 10 )", "PLUS - MINUS") & _
        " As Sales"

cField4 = myiif( _
        " (FLAG = 1 or Flag = 4 )", "PLUS - MINUS") & _
        " As Cash_In"

cField5 = myiif( _
        " (FLAG = 21 AND MINUS < 0)", "MINUS") & _
        " As Bank_In"

cField6 = myiif( _
        " (FLAG = 8)", "PLUS") & _
        " As Trans_In"

cField7 = myiif( _
        " (FLAG = 6)", "PLUS") & _
        " As Income"

cField8 = myiif( _
        " (FLAG = 13)", "PLUS") & _
        " As Chq_In"


' ----------- ’«œ—
cField9 = myiif( _
        " (FLAG = 11 or flag = 12)", "MINUS - PLUS") & _
        " As Purchase"

cField10 = myiif( _
        " (FLAG = 2 or flag = 3)", "MINUS - PLUS") & _
        " As Cash_out"

cField11 = myiif( _
        " (FLAG = 21 AND MINUS > 0)", "MINUS") & _
        " As BANK_OUT"

cField12 = myiif( _
        " (FLAG = 7)", "MINUS") & _
        " As TRANS_OUT"

cField13 = myiif( _
        " (FLAG = 5)", "MINUS") & _
        " As CHARGES"
        
cField14 = myiif( _
        "", "PLUS - MINUS") & _
        " As BalanceLastDay"

' ·÷»ÿ «Œ— Õﬁ·
    cwhere = ""
    If IsDate(XDATE2.Text) Then
        cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
    End If
    
    If XBOX.BoundText <> "" Then
        cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
    End If
            
    cField15 = "(" & _
               "Select Sum(PLUS - MINUS) From BoxMove " & _
               turn(cwhere) & cwhere & _
               ") as Last_Balance"
' «‰ Â«¡ «Œ— Õﬁ·

cString = "Select Date," & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & _
           cField6 & "," & cField7 & "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & "," & cField15 & _
           " From boxmove "
cwhere = ""
If IsDate(xdate1.Text) Then
    cwhere = cwhere & " DATE >= " & DateSq(xdate1.Text)
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If

If IsDate(XDATE2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)
    aHeader(1) = BetweenString(xdate1.Text, XDATE2.Text)
End If

If XBOX.BoundText <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & " box = " & MyParn(XBOX.BoundText)
    aHeader(0) = "—’Ìœ «·Œ“‰… : " & XBOX.Text
End If


cString = cString & turn(cwhere) & cwhere & " Group by Date ORDER BY DATE"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
    nBalance = Val(sourcetable!First_Balance & "") + Val(sourcetable!First_Bal & "")
    nFirst_Balance = nBalance
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = "≈Ã„«·Ì Õ—ﬂ… ÌÊ„Ì… " & XBOX.Text
    temptable!val1 = nBalance
    temptable!VAL14 = nFirst_Balance
    
    temptable!val2 = sourcetable!sales
    temptable!val3 = sourcetable!Cash_in
    temptable!val4 = sourcetable!Bank_in
    temptable!val5 = sourcetable!Trans_In
    temptable!Val6 = sourcetable!Income
    temptable!Val15 = sourcetable!Chq_in
    
    
    temptable!Val7 = sourcetable!purchase
    temptable!Val8 = sourcetable!Cash_out
    temptable!val9 = sourcetable!Bank_Out
    temptable!val10 = sourcetable!Trans_Out
    temptable!val11 = sourcetable!CHARGES
    temptable!val12 = sourcetable!BalanceLastDay + nBalance
    
    temptable!val13 = Val(sourcetable!Last_Balance & "")
    nBalance = Val(temptable!val12 & "")
    
    temptable!Date1 = sourcetable!Date
    temptable!Val21 = nRecord
    temptable!str21 = retHeader(aHeader, 0, 2)
    nRecord = nRecord + 1
    temptable.Update
    sourcetable.MoveNext
Loop

contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\BALBOX2.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
Set sourcetable = Nothing
temptable.Close
Set temptable = Nothing
End Sub
Private Sub Fixgrd()
With grid1
    .Rows = 1
    .Cols = 5
    .TextMatrix(0, 0) = "«· «—ÌŒ"
    .TextMatrix(0, 1) = "≈Ã„«·Ì ’«œ—"
    .TextMatrix(0, 2) = "≈Ã„«·Ì Ê«—œ"
    .TextMatrix(0, 3) = "«·’«›Ì"
    .TextMatrix(0, 4) = "—’Ìœ «·ÌÊ„"
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
End With
End Sub
