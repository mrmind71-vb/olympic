VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salesfrm2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15195
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   600
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   0
      Width           =   3480
      Begin VB.CommandButton cmdTransTo 
         Caption         =   " ÕÊÌ· «·Ì «·»«∆⁄"
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
      End
      Begin VB.CommandButton cmdTransFrom 
         Caption         =   "”Õ» „‰ «·»«∆⁄"
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
         Left            =   1935
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   -945
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   -45
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   4020
      Begin VB.CommandButton cmdOpen 
         Caption         =   "»Ê‰«  „› ÊÕ…"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "„»Ì⁄«  «·ÌÊ„"
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
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… ›« Ê—…"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   5490
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4095
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   540
      Width           =   13425
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8820
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3345
      End
      Begin VB.CommandButton cmdSuper 
         Caption         =   "..."
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xcardesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4185
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   945
         Width           =   2445
      End
      Begin VB.TextBox xCarNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4185
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   585
         Width           =   2445
      End
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄„Ì· ‰ﬁœÌ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11070
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   500
         Width           =   1095
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11070
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo xMan 
         Height          =   360
         Left            =   8820
         TabIndex        =   4
         Top             =   855
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xSuper 
         Height          =   360
         Left            =   4185
         TabIndex        =   10
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊœÌ· «·”Ì«—… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "—’Ìœ «·’‰› :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "«·„‘—› :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·”Ì«—… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
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
         TabIndex        =   8
         Top             =   1260
         Width           =   2445
      End
      Begin VB.Label xbalanceitem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1260
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·Êﬁ  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2595
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1350
         Width           =   660
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12210
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2595
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ«‘Ì— :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12210
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   945
         Width           =   690
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1305
         Width           =   2445
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2595
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2595
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12210
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1215
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         CausesValidation=   0   'False
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   585
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -270
      Top             =   810
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -585
      Top             =   585
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -360
      Top             =   675
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   765
      Top             =   810
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8820
      Width           =   2040
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Move Last"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   555
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   465
      End
   End
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   -90
      Top             =   900
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
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   270
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CheckBox chkprint 
      Alignment       =   1  'Right Justify
      Caption         =   "«·€«¡ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   9540
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   225
      TabIndex        =   44
      Top             =   9495
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Left            =   4275
      TabIndex        =   9
      Top             =   2295
      Width           =   10860
      _cx             =   19156
      _cy             =   11456
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
   Begin VB.Frame Frame7 
      Height          =   960
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   8775
      Width           =   7575
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1035
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox xRateDis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5175
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox xTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2655
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "¬Ã· :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   585
         Width           =   630
      End
      Begin VB.Label xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         TabIndex        =   82
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label xCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
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
         TabIndex        =   81
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label17 
         Caption         =   "‰ﬁœÌ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label14 
         Caption         =   "’«›Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label xDisItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label xTotalDis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label lblTotalQuant 
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label xTotalDisItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   540
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "’«›Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   855
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   165
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "Œ’„ ‰ﬁœÌ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "«·«Ã„«·Ì :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "÷—«∆» «·„»Ì⁄«  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6540
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   2250
      Width           =   4020
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " — Ì» «·„Ã„Ê⁄« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   5940
         Width           =   2040
      End
      Begin VB.TextBox xGroupName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   5895
         Width           =   1680
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«ŸÂ«— «·ﬂ·"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   6345
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VSFlex7LCtl.VSFlexGrid grdGroup 
         Height          =   5640
         Left            =   90
         TabIndex        =   68
         Top             =   180
         Width           =   3840
         _cx             =   6773
         _cy             =   9948
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
         GridColor       =   -2147483633
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
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
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
         AutoResize      =   -1  'True
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
         RightToLeft     =   0   'False
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   8775
      Width           =   2265
      Begin VB.CommandButton cmdVisa 
         Caption         =   "›Ì“«"
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdCash 
         Caption         =   "‰ﬁœÌ"
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
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   675
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox xVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3285
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·„œ›Ê⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   225
         Width           =   765
      End
      Begin VB.Label xPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·»«ﬁÌ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1395
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   585
         Width           =   675
      End
   End
   Begin VB.Label xTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "123456"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   870
      Left            =   4680
      TabIndex        =   65
      Top             =   8865
      Width           =   2805
   End
End
Attribute VB_Name = "salesfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_No As String, lSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String, rdPaid As New ADODB.Recordset
Public bRetvalue As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim Search31 As New Search3, search32 As New Search3, bMarket As Boolean
Public bEdit As Boolean
Dim cFile As String, cFileClient, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
ItemsLookupAll Me, Search3, 10
End Sub
Private Function myreplace() As Boolean
Dim cSaveMode As String
Dim aInsert(16, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(xCode.Text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.Text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "[Notes]"
aInsert(4, 1) = addstring(xNotes.Text)

aInsert(5, 0) = "Discount"
aInsert(5, 1) = Val(xDiscount.Text)

aInsert(6, 0) = "Tax"
aInsert(6, 1) = Val(xTax.Text)

aInsert(7, 0) = "Cash"
aInsert(7, 1) = Val(IIf(Val(xPay.Caption) > Val(xTotal.Caption), xTotal.Caption, xPay.Caption))

aInsert(8, 0) = "Visa"
aInsert(8, 1) = Val(xVisa.Text)

aInsert(9, 0) = "Rate"
aInsert(9, 1) = Val(xRateDis.Text)

aInsert(10, 0) = "userName"
aInsert(10, 1) = addstring(sUserName)

aInsert(11, 0) = "Box"
aInsert(11, 1) = addstring(xBox.BoundText)

aInsert(12, 0) = "Pay"
aInsert(12, 1) = Val(xPay.Caption)

aInsert(13, 0) = "MAN"
aInsert(13, 1) = addstring(xMan.BoundText)

aInsert(14, 0) = "CARNO"
aInsert(14, 1) = addstring(xCarNo.Text)

aInsert(15, 0) = "CARDESCA"
aInsert(15, 1) = addstring(xcardesca.Text)

aInsert(16, 0) = "SUPER"
aInsert(16, 1) = addstring(xSuper.BoundText)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myReplacegrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
xDoc_No.Tag = cSaveMode
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    grid1.EditText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    grid1.TextMatrix(grid1.Row, 3) = "1"
    GrdDesc grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = "1"
        grid1.AddItem ""
        MakeSerial
        grid1_AfterEdit grid1.Row, grid1.col
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.TextMatrix(grid1.Row, 3) = "1"
        grid1_AfterEdit grid1.Row, grid1.col
        grid1.Select grid1.Row + 1, 1
    End If
    CalcTotals
    
ElseIf ActiveControl.Name = CmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If ActiveControl.Name = CmdInform Then Search31.Hide Else Unload Search31
    myload
ElseIf ActiveControl.Name = xCode.Name Then
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    xCode_LostFocus
    Unload search32
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub myproc3(itemCode, itemDesca)
'    If nFound <> -1 Then
'        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    End If
    grid1.EditText = itemCode
    grid1.TextMatrix(grid1.Row, 1) = itemCode
    grid1.TextMatrix(grid1.Row, 2) = mySplit(itemDesca, 1, "(")
    grid1.TextMatrix(grid1.Row, 3) = "1"
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = "1"
        grid1.AddItem ""
        MakeSerial
        grid1_AfterEdit grid1.Row, grid1.col
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.TextMatrix(grid1.Row, 3) = "1"
        grid1_AfterEdit grid1.Row, grid1.col
        grid1.Select grid1.Row + 1, 1
    End If
    CalcTotals
End Sub

Private Sub cmdClient_Click()
publicFlag = 2
Clients.Show 1
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Enabled = True
Else
    Check1.Enabled = False
    Check1.Value = 0
End If
myLoadGroup
End Sub

Private Sub chkprint_Click()
addSetting "print", chkprint.Value, App.Path & "\other.txt"
End Sub

Private Sub cmdCash_Click()
xVisa.Text = ""
xLate.Caption = ""
xCash.Caption = TurnValue(Val(xTotalDis.Caption), 0, "")
xCash_LostFocus
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    'on Error GoTo MyError
    con.BeginTrans
    ' Õ–› «·„” ‰œ
   con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
   con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    
          
    con.CommitTrans
    CardTable.Requery
    
    CmdNewInv_Click
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
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
Private Sub CmdNewInv_Click()
bAddnew = True
mydefine
On Error Resume Next
'xCode.SetFocus
grid1.SetFocus
End Sub
Private Sub cmdSave_Click()
'If myPublic = 0 Then If Not nofoundOther Then Exit Sub
If Not myvalid Then Exit Sub
CashReg.Show 1
If Not lSave Then Exit Sub
mysave
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub

Private Sub cmdSection_Click()

End Sub

Private Sub cmdSuper_Click()
ReDim aPublic(5)
aPublic(0) = "SUPER"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ"
aPublic(4) = "»Ì«‰"
aPublic(5) = "«·„‘—›Ì‰"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
data5.Refresh
End Sub

Private Sub cmdTransFrom_Click()
Dim cString As String
transManfrm.sDate = xDate.Text
transManfrm.sCaption = "”Õ» „‰ Œ“Ì‰… " & xBox.Text
transManfrm.sBox1 = xBox.BoundText
transManfrm.Sbox2 = GetDesca("Select code from file0_50 where type = 1")
transManfrm.Show 1
End Sub

Private Sub cmdTransTo_Click()
Dim cString As String
transManfrm.sDate = xDate.Text
transManfrm.sCaption = "«Ìœ«⁄ ›Ì Œ“Ì‰… " & xBox.Text
transManfrm.Sbox2 = xBox.BoundText
transManfrm.sBox1 = GetDesca("Select code from file0_50 where type = 2")
transManfrm.Show 1
End Sub

Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.BOF And CardTable.EOF Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub cmdItem_Click()
Dim bEditLocal As Boolean
bEditLocal = bEdit: bEdit = True
Items.Show 1
bEdit = bEditLocal
End Sub

Private Sub cmdVisa_Click()
xCash.Caption = ""
xLate.Caption = ""
xVisa.Text = TurnValue(Val(xTotalDis.Caption), 0, "")
xVisa_LostFocus
End Sub

Private Sub cmdopen_Click()
CardLookup "PRINTED = 0"
End Sub

Private Sub Command5_Click()

End Sub

Private Sub fixrate_Click()
Dim SalTable As New ADODB.Recordset
If myPublic = 0 Then
    SalTable.Open "SELECT r_SAL.DOC_ID, r_SAL.RATEDISC, FILE6_20H.rate FROM r_SAL RIGHT JOIN FILE6_20H ON r_SAL.DOC_ID = FILE6_20H.DOC_NO", con, adOpenKeyset, adLockOptimistic, adCmdTableDirect
Else
    SalTable.Open "SELECT r_ret.DOC_ID, r_ret.RATEDISC, FILE6_10H.rate FROM r_ret RIGHT JOIN FILE6_10H ON r_ret.DOC_ID = FILE6_10H.DOC_NO", con, adOpenKeyset, adLockOptimistic, adCmdTableDirect
End If
With SalTable
    .MoveFirst
    Do While Not .EOF
        If myPublic = 0 Then
            con.Execute " UPDATE FILE6_20H SET RATE = " & Val(!RATEDISC & "") & " WHERE DOC_NO = " & MyParn(!doc_ID)
        Else
            con.Execute " UPDATE FILE6_10H SET RATE = " & Val(!RATEDISC & "") & " WHERE DOC_NO = " & MyParn(!doc_ID)
        End If
        Me.Caption = !doc_ID & ""
        .MoveNext
    Loop
End With
End Sub

Private Sub Command1_Click()
doprint
End Sub

Private Sub Command2_Click()
TDaySal.Show 1
End Sub

Private Sub Command3_Click()
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset
loctable.Open "Select * FROM FILE6_20H", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    If Not IsNull(loctable!Time) Then
        dDate = Format(IIf(Val(Format(loctable!Time, "hh")) > 4, loctable!Time, DateAdd("d", -1, loctable!Time)), "dd-mm-yyyy")
        cString = "update file6_20h set file6_20h.date = " & DateSq(dDate)
        cString = cString & turn(cString) & " doc_no = " & MyParn(loctable!doc_no)
        con.Execute cString
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing


Set loctable = New ADODB.Recordset
loctable.Open "Select DOC_NO,SUM(PRICE * QUANT) AS TOTAL FROM FILE6_20 GROUP BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.CASH = " & Val(loctable!Total & "") & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop
MsgBox "done..."
End Sub
Private Sub Form_Activate()
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If Shift = 2 And KeyCode = 83 And xPrinted.Value = 0 and bedit  Then
If xPrinted.Value = 1 Or bEdit = False Then Exit Sub
If KeyCode = 116 Then
    grid1_Validate False
    cmdSave_Click
    KeyCode = 0
ElseIf KeyCode = 115 Then
    itemsgrdfrm.Show 1
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
chkprint.Value = Val(RetSetting("print", App.Path & "\other.txt"))
openCon con
cFileClient = "File3_10"
cdefman = GetDesca("Select code from file6_25 where usercode = " & Val(nusercode))
If cdefman <> "" Then
    cDefBox = GetDesca("Select Code From file0_50 where man = " & cdefman)
End If

'If DefGet(Me.Name, "CUST") = "TRUE" Then
'    cDefClient = GetDesca("Select Min(Code) From file3_10")
'Else
Dim aret As Variant
aret = aGetDesca("select code,desca from file3_10 where cash = 1 order by code")
If UBound(aret) > 0 Then
    cDefClient = aret(1)
    cDefClientDesca = aret(2) & ""
End If
cDefStore = GetDesca("Select code from file0_40 order by code")
bEdit = cdefman <> ""
Select Case myPublic
Case 0
    cCodeDesca = "«·⁄„Ì·"
    cFile = "File6_20"
    cFileHeader = "File6_20H"
    cMoveName = "„»Ì⁄« "
    Me.Caption = "›« Ê—… „»Ì⁄« "
Case 1
    cCodeDesca = "«·⁄„Ì·"
    cFile = "FILE6_10"
    cFileHeader = "File6_10H"
    cMoveName = "„—œÊœ „»Ì⁄« "
    lblClient.Caption = "«·⁄„Ì· :"
    Me.Caption = "›« Ê—… „—œÊœ „»Ì⁄« "
End Select

myLoadGroup

Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT " & cFileHeader & ".*,FILE3_10.DESCA AS CLIENTDESCA,FILE3_10.CASH FROM " & cFileHeader & _
               " LEFT JOIN FILE3_10 ON " & cFileHeader & ".Code = FILE3_10.CODE " & _
               " ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "FILE6_25"
Set xMan.RowSource = data2
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM FILE0_50"
Set xBox.RowSource = data4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

data5.ConnectionString = strCon
data5.RecordSource = "SELECT * FROM SUPER"
Set xSuper.RowSource = data5
xSuper.ListField = "Desca"
xSuper.BoundColumn = "Code"

With grid1
    .Cols = 9
    .Rows = 2
    .Editable = flexEDKbdMouse
End With

Set grid1.DataSource = DATA3
DATA3.ConnectionString = strCon

If sDoc_No <> "" And Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "doc_no = " & MyParn(sDoc_No), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If

If Not (CardTable.EOF And CardTable.BOF) Then
    CmdNewInv_Click
Else
    mydefine
    FixGrd
    xDoc_No.Text = RetZero("1", 6)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SetKbLayout Lang_AR

CardTable.Close
Set CardTable = Nothing

closeCon con

Unload Search3
Unload Search31
Unload search32
Set salesfrm = Nothing
Err.Clear
End Sub

Private Sub grdGroup_DblClick()
If xPrinted.Value = 1 Or bEdit = False Then Exit Sub
grid1.Row = grid1.Rows - 1
grid1.col = 1
listitemsfrm2.nGroup = grdGroup.TextMatrix(grdGroup.Row, 1)
listitemsfrm2.cGroupname = grdGroup.TextMatrix(grdGroup.Row, 0)
listitemsfrm2.Show 1
End Sub

Private Sub grdGroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then grdGroup_DblClick
End Sub

Private Sub grdGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 0 Then Exit Sub
If Button <> 1 Then Exit Sub
With grdGroup
Dim nRow
If .Rows = 0 Then Exit Sub
If r = -1 Then Exit Sub
r = .Row
r = .DragRow(r)
ReplaceGroup
End With
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
myreplaceRow Row, col
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    xBalance.Caption = ""
    If Not validRows(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
grid1.EditMaxLength = IIf(col = 3, 6, 0)
End Sub

Private Sub grid1_EnterCell()
If xPrinted.Value = 1 Or bEdit = False Then
    grid1.Editable = flexEDNone
ElseIf (grid1.col = 1 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "") Or grid1.col = 2 Or grid1.col = 4 Or grid1.col = 5 Or grid1.col = 7 Or grid1.col = 8 Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 1 Then
    grid1.Editable = flexEDKbdMouse
ElseIf grid1.col = 1 Then
     grid1.Editable = flexEDKbdMouse
Else
   grid1.Editable = IIf(Trim(grid1.TextMatrix(grid1.Row, 1)) <> "", flexEDKbdMouse, flexEDNone)
   'Grid1.EditCell
End If
If grid1.col = 3 Then
    xBalance.Caption = LastBalance(grid1.TextMatrix(grid1.Row, 1), xStore.BoundText, con)
End If
With grid1
'    If .Col = 6 And .Row > 1 And (.TextMatrix(.Row, 1) <> "") Then
'        If .TextMatrix(.Row, 6) = "" Then .TextMatrix(.Row, 6) = .TextMatrix(.Row - 1, 6)
'    End If
End With

End Sub
Private Sub Grid1_GotFocus()
If grid1.Rows < 2 Then Exit Sub
If grid1.Row = 0 Then
    grid1.Row = 1
    grid1.col = 1
End If
grid1_EnterCell
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    myPos
End If

If KeyCode = 115 Or (KeyCode = 13 And Shift = 2) Then xDiscount.SetFocus
'If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then
'    If Not validRows(, False, True) Then Exit Sub
'    grid1.AddItem "", grid1.Row
'    MakeSerial
'    bInsert = True
'End If
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRows(grid1.Row) Then
    grid1.AddItem "", grid1.Row
End If
If KeyCode = 112 And xPrinted.Value = 0 And bEdit = True Then
    grid1.Row = grid1.Rows - 1
    grid1.col = 1
    ItemsLookup
End If
End Sub

Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    myPos
End If
End Sub
Private Sub myPos()
    If grid1.col = 3 Then
        grid1.Row = grid1.Row + 1
        grid1.col = IIf(grid1.Row = grid1.Rows - 1, 1, 3)
   ElseIf grid1.col = 1 Then
        grid1.col = 3
     End If
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)

'If KeyAscii = 13 Then
'    Select Case Col
'        Case 1
'            If Row = grid1.Rows - 2 Then
'                grid1.TextMatrix(grid1.Rows - 1, 3) = ""
'                grid1.AddItem ""
'                grid1.Select grid1.Rows - 2, 3
'                MakeSerial
'            Else
'                grid1.TextMatrix(grid1.Row, 3) = ""
'                grid1.Select Row, 3
'            End If
'        Case 3
'            grid1.Select iif(row = grid1.Rows -2,, grid1.Rows - 2, 6
'        Case 6
'            grid1.Select grid1.Rows - 1, 1
'    End Select
    'CalcTotals
'End If
End Sub
Private Sub grid1_LostFocus()
SetKbLayout Lang_AR
End Sub


Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
'If Col = 3 And grid1.TextMatrix(Row, 1) <> "" And xstore.BoundText <> "" Then
'    nBalance = LastBalance(grid1.TextMatrix(Row, 1), xstore.BoundText, CON)
'    xbalanceitem.Caption = nBalance
'End If
If col = 3 Then xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
If Not validRows(grid1.Row) And grid1.Row <> grid1.Rows - 1 Then grid1.RemoveItem grid1.Row
xBalance.Caption = ""
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If grid1.EditText = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    Else
        If GetDesca("select item from file1_10 where item = " & MyParn(grid1.EditText)) = "" Then
            Dim cAlter As String
            cAlter = GetDesca("Select item from item_alter where itemsub = " & MyParn(grid1.EditText))
            If cAlter <> "" Then
                grid1.EditText = cAlter
            Else
                MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End If

If col = 3 Or col = 4 Then
   If Not IsNumeric(grid1.EditText) Then
        Cancel = True
        Exit Sub
   End If
End If
'If Row = grid1.Rows - 1 And Cancel = False Then
'    grid1.AddItem ""
'End If
End Sub

Private Sub XBOX_Click(Area As Integer)
'If Not xDoc_No.Enabled Then UpdateHeader
End Sub
Private Sub xBox_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub
Private Sub xBox_LostFocus()
xCash.Enabled = (Trim(xBox.BoundText) <> "")
cmdCash.Enabled = (Trim(xBox.BoundText) <> "")
CalcTotals
'If Not xBox.MatchedWithList Then
'    xBox.BoundText = ""
'    xCash.Caption = ""
'    xLate.Caption = Val(xTotalDis.Caption) - Val(xVisa.Text)
'End If
'If Not xDoc_No.Enabled Then UpdateHeader
xBox.BackColor = &H80000005
End Sub
Private Sub xCash_LostFocus()
'CalcLate xCash
'CalcTotals
'If Not xDoc_No.Enabled Then UpdateHeader
End Sub

Private Sub xCode_DblClick()
CLIENTLOOKUP
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
Dim aret
aret = aGetDesca("select desca,cash from " & cFileClient & " where code = " & MyParn(xCode.Text))
If UBound(aret) > 0 Then
    xCodeDesca.Caption = aret(1) & ""
    chkCash.Value = IIf(aret(2), 1, 0)
End If
End Sub

Private Sub xCode_Validate(Cancel As Boolean)
If Trim(xCode.Text) = "" Then Cancel = True
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.Text) Then Cancel = True
End Sub

Private Sub xDiscount_LostFocus()
xDiscount.BackColor = &H80000005
CalcTotals
'UpdateHeader
End Sub
Private Function myvalid() As Boolean
Dim i As Integer
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
'If Not validRows(, False, True) Then
    'MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
'    Exit Function
'End If
'DelValid , True
End With
myvalid = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
On Error GoTo myerror
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xBox.BoundText = CardTable!BOX & ""
xMan.BoundText = CardTable!MAN & ""
xNotes.Text = CardTable!Notes & ""
xCarNo.Text = CardTable!carno & ""
xcardesca.Text = CardTable!carDesca & ""
xSuper.BoundText = CardTable!super & ""
xCode.Text = CardTable!Code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
chkCash.Value = IIf(CardTable!CASH, 1, 0)
'xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = TurnValue(Val(CardTable!discount & ""), 0, "")
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")
xCash.Caption = TurnValue(Val(CardTable!CASH & ""), 0, "")
xVisa.Text = TurnValue(Val(CardTable!Visa & ""), 0, "")
xPay.Caption = Myvalue(Val(CardTable!Pay & ""))
xPrinted.Value = IIf(CardTable!printed, 1, 0)
xTotal.ForeColor = IIf(xPrinted.Value = 0, &H80&, &H808080)
xtime.Caption = Format(CardTable!Time, "hh:nn")
'If Not bLeaveBal Then xBalance.Caption = ""
myloadgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
xDate.Text = Format(IIf(Val(Format(Time, "hh")) > 4, Date, DateAdd("d", -1, Date)), "dd-mm-yyyy")
xBalance.Caption = ""
xBox.BoundText = cDefBox
xCode.Text = cDefClient
xCarNo.Text = ""
xcardesca.Text = ""
xSuper.BoundText = ""
xCodeDesca.Caption = cDefClientDesca
xStore.BoundText = cDefStore
xDiscount.Text = ""
chkCash.Value = 1
xTotalDisItem.Caption = ""
xDisItem.Caption = ""
xTotal.Caption = ""
xTax.Text = ""
xLate.Caption = ""
xVisa.Text = ""
xCash.Caption = ""
xPrinted.Value = 0
'xbalanceitem.Caption = ""
xMan.BoundText = cdefman
xTotalItem.Caption = ""
xTotalDis.Caption = ""
xNotes.Text = ""
xtotalQuant.Caption = ""
xRate.Text = ""
xRest.Caption = ""
xPay.Caption = ""
xtime.Caption = Format(Time, "hh:nn")
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
FixGrd
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit) And xPrinted.Value = 0 And bEdit
CmdDelInv.Enabled = nMode = LoadMode And bEdit And xPrinted.Value = 0
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = DefineMode
xCash.Enabled = (Trim(xBox.BoundText) <> "")
cmdCash.Enabled = (Trim(xBox.BoundText) <> "")
xDoc_No.Tag = nMode
xCode.Enabled = xPrinted.Value = 0 And bEdit = True
End Sub
Private Function retBool(cFieldName) As Boolean
If Not (CardTable.EOF Or CardTable.BOF) Then
    retBool = CardTable(cFieldName)
End If
End Function
Private Sub xDoc_No_LostFocus()
xDoc_No.BackColor = &H80000005
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And xPrinted.Value = 0 And bEdit And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        CalcTotals
        'UpdateHeader
        MakeSerial grid1.Row
    End If
End If
'If KeyCode = 27 Then xDate.SetFocus
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
'Select Case Grid1.Col
'    Case 1
'        If KeyCode = 27 Then
'            Exit Sub
'        End If
'        If KeyCode = 112 Then
'            ItemsLookup
'        End If
'End Select

End Sub
Private Sub GrdDesc(Row)
grid1.TextMatrix(Row, 7) = ""
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
Dim aret As Variant
aret = aGetDesca("select desca,Price3,Discount,price2,package,unit from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 1)))
If UBound(aret) <> 0 Then
    grid1.TextMatrix(Row, 2) = aret(1) & ""
    grid1.TextMatrix(Row, 3) = 1
    grid1.TextMatrix(Row, 4) = aret(2) & ""
    grid1.TextMatrix(Row, 6) = aret(3) & ""
    xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
    'grid1.TextMatrix(Row, 7) = aRet(5) & ""
    'grid1.TextMatrix(Row, 8) = aRet(6) & ""
End If
CalcTotals
End Sub
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalDis As Double
With grid1
For i = 1 To grid1.Rows - 1
    nTotalitem = nTotalitem + (Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)))
    nDiscount = 1 - (Val(.TextMatrix(i, 6)) / 100)
    grid1.TextMatrix(i, 5) = Myvalue(Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)) * nDiscount)
    nTotalDisItem = nTotalDisItem + (Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)) * nDiscount)
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 3))
Next
nDisItem = nTotalitem - nTotalDisItem
nTotalDis = nTotalDisItem - Val(xDiscount.Text)

nTotal = nTotalDis + Val(xTax.Text)

xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
xTotalItem.Caption = Format(nTotalitem, "Fixed")
xTotalDisItem.Caption = Format(nTotalDisItem, "Fixed")
xDisItem.Caption = Format(nDisItem, "Fixed")
xTotalDis.Caption = Format(nTotalDis, "Fixed")

xTotal.Caption = Format(nTotal, "Fixed")
xtotalQuant.Caption = Format(nTotalQuant, "#0.0000")

xRest.Caption = IIf(Val(xPay.Caption) - Val(xTotal.Caption) < 0, "", Val(xPay.Caption) - Val(xTotal.Caption))
'If nMode = 0 Then
'    xCash.Caption = Format(IIf(Val(xPay.Caption) >= Val(xTotal.Caption), nTotal, Val(xPay.Caption)), "fixed")
'End If


If nMode = 0 Then
    If xBox.BoundText = "" Then
        xCash.Caption = Format(0, "Fixed")
        xLate.Caption = Format(nTotal - Val(xVisa.Text), "Fixed")
    Else
        nCash = IIf(Val(xPay.Caption) >= Val(xTotal.Caption), nTotal, Val(xPay.Caption))
        xCash.Caption = Format(nCash, "Fixed")
        xLate.Caption = nTotal - (Val(xCash.Caption))
    End If
End If

   

'End If

'If Val(xLate.Text) < 0 Then
    
'End If
'If Val(xVisa.Text) < 0 Then
'    xVisa.Text = ""
'    If Val(xLate.Caption) <> 0 Then
'        xLate.Caption = Val(xLate.Caption) - Abs(Val(xVisa.Text))
'    Else
'
'    End If
'End If
End With
End Function
Private Sub CardLookup(Optional pWhere As String = "")
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,DATE , Convert(VARCHAR(10),[DATE],111), " & cFileClient & ".Desca " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE )"
If pWhere <> "" Then
Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & pWhere
End If
Generalarray(2) = "Order by Date , doc_no "
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ " & cCodeDesca
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search31.Caption = "«” ⁄·«„"
Search31.grid1.fontsize = 10
Search31.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow And Trim(grid1.TextMatrix(i, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True


End Sub

Private Sub xGroupName_Change()
Check1.Enabled = Trim(xGroupName.Text) = ""
If Trim(xGroupName.Text) <> "" Then
    Check1.Value = 0
End If
myLoadGroup
End Sub

Private Sub xMAN_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xman_LostFocus()
'If Not xDoc_No.Enabled Then UpdateHeader
xMan.BackColor = &H80000005
End Sub

Private Sub xMAN_Validate(Cancel As Boolean)
If Not xMan.MatchedWithList Then xMan.BoundText = ""
If Trim(xMan.BoundText) = "" Then Cancel = True
End Sub

Private Sub xNotes_LostFocus()
'If Not xDoc_No.Enabled Then UpdateHeader
End Sub

Private Sub xRate_LostFocus()
xRate.BackColor = &H80000005
If Val(xRate.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
'UpdateHeader
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!store Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    'If Not (movetable!Type = cItemmove And movetable!Doc_Id = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    'End If
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub FixGrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ì|" & "«·Œ’„|" & "«·⁄»Ê…|" & "«·ÊÕœ…|"
.ColWidth(0) = 500
.ColWidth(1) = 1800
.ColWidth(2) = 4500
.ColWidth(3) = 1100
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1300
.ColHidden(.Cols - 4) = True
.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load search32
search32.Caption = "«” ⁄·«„"
search32.Show 1
End Sub

Private Sub xRateDis_Lostfocus()
xDiscount.Text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.Text) / 100))
'CalcTotals
'UpdateHeader
End Sub

Private Sub xStore_GotFocus()
ActiveControl.BackColor = &HC0FFFF
End Sub

Private Sub xstore_LostFocus()
xStore.BackColor = &H80000005
'If Not xDoc_No.Enabled Then UpdateHeader
End Sub

Private Sub xStore_Validate(Cancel As Boolean)
If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub

Private Sub xSuper_LostFocus()
If Not xSuper.MatchedWithList Then xSuper.BoundText = ""
End Sub

Private Sub xTax_LostFocus()
xTax.BackColor = &H80000005
CalcTotals
'UpdateHeader
End Sub
Private Function RemoveItem(nRow) As Boolean
On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From " & cFile & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
    For i = nRow + 1 To grid1.Rows - 2
        con.Execute "update " & cFile & " set [row] = " & (i - 1) & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
    Next
End If
grid1.RemoveItem nRow
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function UpdateHeaderOld()
Dim aInsert(11, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(xCode.Text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.Text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "[Notes]"
aInsert(4, 1) = addstring(xNotes.Text)

aInsert(5, 0) = "Discount"
aInsert(5, 1) = Val(xDiscount.Text)

aInsert(6, 0) = "Tax"
aInsert(6, 1) = Val(xTax.Text)

aInsert(7, 0) = "Cash"
aInsert(7, 1) = Val(xCash.Caption)

aInsert(8, 0) = "Visa"
aInsert(8, 1) = Val(xVisa.Text)

aInsert(9, 0) = "Rate"
aInsert(9, 1) = Val(xRateDis.Text)

aInsert(10, 0) = "Total"
aInsert(10, 1) = Val(xTotal.Caption)

aInsert(11, 0) = "userName"
aInsert(11, 1) = addstring(sUserName)

On Error GoTo myerror
con.BeginTrans
con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
con.CommitTrans
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function validHeader(Optional bMsg As Boolean = True) As Boolean
If Trim(xDoc_No.Text) = "" Then
    If bMsg Then MsgBox "—ﬁ„ «·›« Ê—… €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xDate.Text) Then
    If bMsg Then MsgBox "«· «—ÌŒ €Ì— ’«·Õ «Ê „”Ã·"
    Exit Function
End If

If Trim(xStore.BoundText) = "" Then
    If bMsg Then MsgBox "«·„Œ“‰ €Ì— „”Ã·"
    Exit Function
End If

If Trim(xCode.Text) = "" Then
    If bMsg Then MsgBox "ﬂÊœ «·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If

validHeader = True
End Function

Private Sub xVisa_LostFocus()
'If Val(xLate.Caption) = 0 Then
'    Xcash.caption = TurnValue(Val(xTotalDis.Caption) - Val(xVisa.Text), 0, "")
'Else
'    xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(Xcash.caption) - Val(xVisa.Text), 0, "")
'End If
CalcLate xVisa
'CalcTotals
If Not xDoc_No.Enabled Then UpdateHeader
End Sub
Private Sub CalcLate(pControl)
xLate.Caption = TurnValue(Val(xTotalDis.Caption) - Val(xCash.Caption) - Val(xVisa.Text), 0, "")
If Val(xLate.Caption) < 0 Then
    If pControl.Name = xCash.Name Then
        If Val(xVisa.Text) >= Abs(Val(xLate.Caption)) Then
            xVisa.Text = Val(xVisa.Text) - Abs(Val(xLate.Caption))
            xLate.Caption = 0
        Else
            xLate.Caption = Val(xVisa.Text) - Abs(Val(xLate.Caption))
            xVisa.Text = ""
        End If
    Else
        If Val(xCash.Caption) >= Abs(Val(xLate.Caption)) Then
            xCash.Caption = Val(xCash.Caption) - Abs(Val(xLate.Caption))
            xLate.Caption = 0
        Else
            xLate.Caption = Val(xCash.Caption) - Abs(Val(xLate.Caption))
            xCash.Caption = ""
        End If
    End If
End If
End Sub
Private Function validRows(Optional prow = -1, Optional igMsg As Boolean = True, Optional bReqQuant As Boolean = False) As Boolean
For nRow = IIf(prow = -1, 1, prow) To IIf(prow = -1, grid1.Rows - 2, prow)
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then
        If Not igMsg Then MsgBox "«·’‰› ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã· "
        Exit Function
    End If
'    If Val(grid1.TextMatrix(nRow, 3)) = 0 And bReqQuant Then
'        If Not igMsg Then MsgBox "«·ﬂ„Ì… ›Ï «·”ÿ— —ﬁ„ " & nRow & " €Ì— „”Ã·… "
'        Exit Function
'    End If
Next
validRows = True
End Function
Sub myproc2(nDoc_no)
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
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
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem))
End Function
Private Sub doprintALL()
Dim aHeader(2)
If Not myvalid Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim AllDocTable As New ADODB.Recordset
Dim cAllDoc As String
Dim nVal5 As Double, nVal6 As Double, nVal7 As Double, nVal8 As Double, nVal9 As Double
If myPublic = 0 Then
    cStr1 = "SELECT DOC_NO FROM FILE6_20H WHERE FILE6_20H.CODE = " & MyParn(xCode.Text) & " AND FILE6_20H.DATE = " & DateSq(xDate.Text) & " GROUP BY DOC_NO ORDER BY DOC_NO "
Else
    cStr1 = "SELECT DOC_NO FROM FILE6_10H WHERE FILE6_10H.CODE = " & MyParn(xCode.Text) & " AND FILE6_10H.DATE = " & DateSq(xDate.Text) & " GROUP BY DOC_NO ORDER BY DOC_NO "
End If
AllDocTable.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText
With AllDocTable
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    cAllDoc = !doc_no
    .MoveNext
    If .EOF Then cAllDoc = ""
    Do While Not .EOF
        cAllDoc = cAllDoc & "-" & !doc_no
        .MoveNext
    Loop
End With
If myPublic = 0 Then
    cStr1 = "SELECT  MAX(FILE6_20H.DOC_NO) AS M_DOC_NO , FILE6_20.ITEM, FILE1_10.DESCA, FILE6_20.PRICE, SUM(FILE6_20.QUANT) AS TQUANT , FILE6_20.discount , " & _
            " Sum([FILE6_20].[QUANT] * [FILE6_20].[PRICE] *(1-( [FILE6_20].[DISCOUNT]/100))) AS T_TOTAL " & _
            " FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM  " & _
            " WHERE FILE6_20H.CODE = " & MyParn(xCode.Text) & " AND FILE6_20H.DATE = " & DateSq(xDate.Text) & _
            " GROUP BY FILE6_20.ITEM, FILE1_10.DESCA, FILE6_20.PRICE, FILE6_20.discount  "
Else
    cStr1 = "SELECT  MAX(FILE6_10H.DOC_NO) AS M_DOC_NO , FILE6_10.ITEM, FILE1_10.DESCA, FILE6_10.PRICE, SUM(FILE6_10.QUANT) AS TQUANT , FILE6_10.discount,  " & _
            " Sum([FILE6_10].[QUANT] * [FILE6_10].[PRICE] * (1-([FILE6_10].[DISCOUNT]/100))) AS T_TOTAL " & _
            " FROM (FILE6_10 INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO) LEFT JOIN FILE1_10 ON FILE6_10.ITEM = FILE1_10.ITEM  " & _
            " WHERE FILE6_10H.CODE = " & MyParn(xCode.Text) & " AND FILE6_10H.DATE = " & DateSq(xDate.Text) & _
            " GROUP BY FILE6_10.ITEM, FILE1_10.DESCA, FILE6_10.PRICE, FILE6_10.discount "
End If

sourcetable.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
    If myPublic = 0 Then
        NH = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_20H WHERE DATE = " & DateSq(xDate.Text) & " AND CODE = " & MyParn(xCode.Text)) & "")
        nVal5 = Val(GetDesca("SELECT SUM(FILE6_20.QUANT * FILE6_20.PRICE ) FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.Text) & " AND FILE6_20H.DATE = " & DateSq(xDate.Text)) & "")
        nVal6 = Val(GetDesca("SELECT SUM(FILE6_20.QUANT  * FILE6_20.PRICE  * ((FILE6_20.DISCOUNT /100))) FROM FILE6_20  INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.Text) & " AND FILE6_20H.DATE = " & DateSq(xDate.Text)) & "")
        nVal7 = NH
        nVal8 = nVal5 - nVal6 - NH
        nVal9 = Val(GetDesca("SELECT SUM(FILE6_20.QUANT ) FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.Text) & " AND FILE6_20H.DATE = " & DateSq(xDate.Text)) & "")
    Else
        NH = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_10H WHERE DATE = " & DateSq(xDate.Text) & " AND CODE = " & MyParn(xCode.Text)) & "")
        nVal5 = Val(GetDesca("SELECT SUM(FILE6_10.QUANT * FILE6_10.PRICE ) FROM FILE6_10 INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO WHERE FILE6_10H.CODE = " & MyParn(xCode.Text) & " AND FILE6_10H.DATE = " & DateSq(xDate.Text)) & "")
        nVal6 = Val(GetDesca("SELECT SUM(FILE6_10.QUANT  * FILE6_10.PRICE  * ((FILE6_10.DISCOUNT /100))) FROM FILE6_10  INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO WHERE FILE6_10H.CODE = " & MyParn(xCode.Text) & " AND FILE6_10H.DATE = " & DateSq(xDate.Text)) & "")
        nVal7 = NH
        nVal8 = nVal5 - nVal6 - NH
        nVal9 = Val(GetDesca("SELECT SUM(FILE6_10.QUANT ) FROM FILE6_10 INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO WHERE FILE6_10H.CODE = " & MyParn(xCode.Text) & " AND FILE6_10H.DATE = " & DateSq(xDate.Text)) & "")
    End If


With sourcetable
If sourcetable.RecordCount > 0 Then
    Do While Not sourcetable.EOF
        temptable.AddNew
        If myPublic = 0 Then
            temptable!str21 = "≈–‰  ”·Ì„ »÷«⁄… "
        Else
            temptable!str21 = "≈–‰ „— Ã⁄ »÷«⁄… "
        End If
        temptable!str1 = !M_DOC_NO
        temptable!str2 = xDate.Text
        temptable!str3 = Format(xCode.Text)
        temptable!str4 = xCodeDesca.Caption
        
        temptable!str6 = IIf(Val(xLate.Caption) = 0, "‰ﬁœÌ", "¬Ã·")
        temptable!Str11 = !Item
        If cAllDoc <> "" Then temptable!str5 = cAllDoc
        
        temptable!str12 = !Desca
        temptable!val1 = !TQUANT
        temptable!val2 = !price
        If myPublic = 0 Then
            temptable!val3 = ![discount]
        Else
            temptable!val3 = ![discount]
        End If
        temptable!val4 = !T_TOTAL
'       temptable!Val10 = !
        
        temptable!val5 = nVal5
        temptable!Val6 = nVal6
        temptable!Val7 = nVal7
        temptable!Val8 = nVal8
        temptable!val9 = nVal9
        
        temptable!str10 = MyOnly(nVal8)
        temptable!val9 = myPublic
        temptable.Update
        sourcetable.MoveNext
    Loop
End If
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\Tsales.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Function myReplacegrd() As Boolean
Dim aInsert(6, 1)
With grid1
    For i = 1 To grid1.Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 3))

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(i, 4))

        aInsert(4, 0) = "Discount"
        aInsert(4, 1) = Val(.TextMatrix(i, 6))

        aInsert(5, 0) = "Cost"
        aInsert(5, 1) = LastCostDate(grid1.TextMatrix(i, 1), xDate.Text, con)

        aInsert(6, 0) = "row"
        aInsert(6, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myReplacegrd = True
End Function
Private Function myreplaceGrdRow(i) As Boolean
Dim aInsert(5, 1)
With grid1
con.BeginTrans
aInsert(0, 0) = "doc_no"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "item"
aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))

aInsert(2, 0) = "quant"
aInsert(2, 1) = Val(.TextMatrix(i, 3))

aInsert(3, 0) = "Price"
aInsert(3, 1) = Val(.TextMatrix(i, 4))

aInsert(4, 0) = "Discount"
aInsert(4, 1) = Val(.TextMatrix(i, 6))

aInsert(5, 0) = "row"
aInsert(5, 1) = i

If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, cFile)
Else
    con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(i, .Cols - 1))
End If
End With
con.CommitTrans
If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then myloadgrd
myreplaceGrdRow = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myloadgrd()
With grid1
    cField1 = "case when " & cFile & ".Discount = 0 then Null else " & cFile & ".Discount end "
    cString = "SELECT " & cFile & ".ROW, " & cFile & ".ITEM, FILE1_10.DESCA, Quant, " & cFile & ".Price,0 AS Expr1," & cField1 & ",FILE1_10.PACKAGE,FILE1_10.[UNIT],ID " & _
          " FROM " & cFile & " LEFT JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " ORDER by " & cFile & ".ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
    grid1.AddItem ""
    MakeSerial
End With
Handlecontrols LoadMode
CalcTotals
FixGrd
End Sub
Private Sub UpdateHeader()
If Not validHeader Then Exit Sub
myreplace
CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub xusername_GotFocus()
xusername.SelStart = 0
xusername.SelLength = Len(xusername.Text)
End Sub
Private Sub xNotes_GotFocus()
xNotes.SelStart = 0
xNotes.SelLength = Len(xNotes.Text)
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
End Sub
Private Sub xDoc_No_GotFocus()
xDoc_No.SelStart = 0
xDoc_No.SelLength = Len(xDoc_No.Text)
End Sub
Private Sub xdate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xRate_GotFocus()
xRate.SelStart = 0
xRate.SelLength = Len(xRate.Text)
End Sub
Private Sub xRateDis_GotFocus()
xRateDis.SelStart = 0
xRateDis.SelLength = Len(xRateDis.Text)
End Sub
Private Sub xDiscount_GotFocus()
xDiscount.SelStart = 0
xDiscount.SelLength = Len(xDiscount.Text)
End Sub
Private Sub xTax_GotFocus()
xTax.SelStart = 0
xTax.SelLength = Len(xTax.Text)
End Sub
Private Sub xVisa_GotFocus()
xVisa.SelStart = 0
xVisa.SelLength = Len(xVisa.Text)
End Sub
Private Sub xCarno_GotFocus()
With xCarNo
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub xCardesca_GotFocus()
With xcardesca
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
lSave = False
If Not myvalid Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
If bEnd Then
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
    If bPrint Then
        If chkprint.Value = 0 Then
            If doprint Then SavePrint
        Else
            openCash
            SavePrint
        End If
    End If
    If Val(xRest.Caption) <> 0 Then InformOk " «·»«ﬁÌ " & xRest.Caption
    mydefine
Else
    CardTable.Requery
    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Function
Private Sub myLoadGroup()
With grdGroup
grdGroup.ColHidden(1) = True
grdGroup.ColWidth(0) = .Width - 400
.ColAlignment(0) = flexAlignCenterCenter
grdGroup.Rows = 0
Dim loctable As New ADODB.Recordset
cString = "select FILE1_50.CODE,FILE1_50.DESCA from FILE1_50 inner join file1_10 on file1_50.code = file1_10.[GROUP] "
'cString = cString & turn(cString) & " FILE1_50.SHOW = 1"
If Check2.Value = 0 Then
    'cString = cString & turn(cString) & "(NOT FILE1_10.ITEM2 IS NULL)"
    cString = cString & turn(cString) & "(FILE1_10.SHOW = 1)"
End If
If Trim(xGroupName.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xGroupName.Text, "FILE1_50.desca")
End If
cString = cString & " GROUP BY FILE1_50.CODE,FILE1_50.DESCA,FILE1_50.SERIAL order by serial"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    grdGroup.AddItem ""
    grdGroup.TextMatrix(grdGroup.Rows - 1, 0) = loctable!Desca & ""
    grdGroup.TextMatrix(grdGroup.Rows - 1, 1) = loctable!Code & ""
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End With
End Sub
Private Function ReplaceGroup() As Boolean
With grdGroup
On Error GoTo myerror
con.BeginTrans
For i = 0 To .Rows - 1
    cString = "UPDATE FILE1_50 SET FILE1_50.SERIAL = " & i & _
               " WHERE CODE = " & .TextMatrix(i, .Cols - 1)
    con.Execute cString
Next
con.CommitTrans
ReplaceGroup = True
End With
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function doprint() As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not myvalid Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!Date1 = DateFix(xDate.Text)
    
    temptable!str5 = TurnValue(xtime.Caption)
    If xNotes.Text <> "" Then temptable!str2 = "«·”«œ… : " & xNotes.Text
    temptable!Str11 = "ÿ—Ìﬁ 14 „«ÌÊ - «„«„ „—ﬂ“ «”ﬂ‰œ—Ì…"
    
    temptable!str12 = TurnValue(turn(xSuper.Text, "«·„‘—› : ") & xSuper.Text)
    temptable!Str13 = TurnValue(turn(xCarNo.Text, "—ﬁ„ «·”Ì«—… : ") & xCarNo.Text)
    temptable!Str14 = TurnValue(turn(xcardesca.Text, "„ÊœÌ· : ") & xcardesca.Text)
    
    temptable!str3 = ArbString(Val(xDoc_No.Text))
    temptable!str6 = .TextMatrix(i, 2)
    temptable!str4 = TurnValue(xMan.Text)
    temptable!val1 = Val(.TextMatrix(i, 3))
    temptable!val2 = Val(.TextMatrix(i, 4))
    temptable!val3 = Val(.TextMatrix(i, 5))
    temptable!Str11 = ArbString("ÿ—Ìﬁ 14 „«ÌÊ - «„«„ «·„—ﬂ“ «·ÿ»Ì")
    temptable!val4 = Val(xTotalItem.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xCash.Caption)
    temptable!Val7 = Val(xPay.Caption)
    temptable!Val8 = Val(xRest.Caption)
    temptable.Update
Next i
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.Destination = crptToPrinter
main.REPORT1.ReportFileName = App.Path & "\Reports\sales_bon.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
main.REPORT1.Destination = crptToWindow
doprint = True
GoTo closeCon
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon:
temptable.Close
Set temptable = Nothing
End Function
Private Sub SavePrint()
On Error GoTo myerror
con.BeginTrans
con.Execute "update file6_20h set FILE6_20H.PRINTED = 1 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myreplaceRow(Row, col, Optional pLookup As Boolean = False)
Dim nBalance As Double

If Not validRows(Row) Then Exit Sub

If Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    MakeSerial
End If

If col = 1 Then
    GrdDesc Row
End If


CalcTotals


If xDoc_No.Tag = DefineMode Or (Row <> grid1.Rows - 2 And grid1.TextMatrix(Row, grid1.Cols - 1) = "") Then
    mysave False
Else
    myreplaceGrdRow Row
End If

End Sub
Private Function doprint2() As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not myvalid Then Exit Function
Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.Destination = crptToPrinter
main.REPORT1.ReportFileName = App.Path & "\Reports\sales_bon.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
main.REPORT1.Destination = crptToPrinter
doprint2 = True
GoTo closeCon
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
closeCon:
End Function
