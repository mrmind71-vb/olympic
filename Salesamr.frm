VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form salesfrmamr 
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
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Salesamr.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Salesamr.frx":241E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2775
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Salesamr.frx":4CB8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4140
         Picture         =   "Salesamr.frx":7264
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.Frame Frame8 
      Height          =   690
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   8550
      Width           =   2310
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1620
         Picture         =   "Salesamr.frx":9A37
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "«·«ŒÌ—"
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   600
         Picture         =   "Salesamr.frx":BCCE
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         Picture         =   "Salesamr.frx":DEFB
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "«·«Ê·"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1110
         Picture         =   "Salesamr.frx":101D8
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1140
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   1350
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Salesamr.frx":12424
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Salesamr.frx":1499D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CommandButton cmdTransTo 
         Caption         =   " ÕÊÌ· «·Ì «·»«∆⁄"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   135
         Width           =   1500
      End
      Begin VB.CommandButton cmdTransFrom 
         Caption         =   "”Õ» „‰ «·»«∆⁄"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1575
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   -675
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   -270
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   3930
      Begin VB.CommandButton cmdOpen 
         Caption         =   "»Ê‰«  „› ÊÕ…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2565
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
         Caption         =   "„»Ì⁄«  «·ÌÊ„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… ›« Ê—…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   675
      Width           =   13425
      Begin VB.TextBox xDoc_No 
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
         Height          =   345
         Left            =   11070
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11070
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   1095
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8820
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3345
      End
      Begin VB.CheckBox chkBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷"
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
         Height          =   195
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄„Ì· ‰ﬁœÌ"
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   9405
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
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
         Top             =   135
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   345
         Left            =   90
         TabIndex        =   3
         Top             =   495
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSDataListLib.DataCombo xBox 
         Height          =   345
         Left            =   90
         TabIndex        =   5
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
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
      Begin MSDataListLib.DataCombo xMan 
         Height          =   345
         Left            =   8820
         TabIndex        =   4
         Top             =   855
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "—’Ìœ «·’‰› :"
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1305
         Width           =   1170
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1305
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
         TabIndex        =   44
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
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1395
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ«‘Ì— :"
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
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   990
         Width           =   615
      End
      Begin VB.Label xBalance 
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
         Height          =   360
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label xCodeDesca 
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
         Height          =   345
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   2625
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   855
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
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
         Left            =   12240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   615
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
      TabIndex        =   49
      Top             =   270
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CheckBox chkprint 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·€«¡ «·ÿ»«⁄…"
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
      Height          =   195
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   8190
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   450
      TabIndex        =   27
      Top             =   9270
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Height          =   960
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8550
      Width           =   2130
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·„œ›Ê⁄ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
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
         TabIndex        =   46
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
         TabIndex        =   38
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·»«ﬁÌ :"
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
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   585
         Width           =   675
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6000
      Left            =   4275
      TabIndex        =   8
      Top             =   2520
      Width           =   10815
      _cx             =   19076
      _cy             =   10583
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
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame4 
      Height          =   5640
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2475
      Width           =   3975
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " — Ì» «·„Ã„Ê⁄« "
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
         Height          =   285
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   5220
         Width           =   1815
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
         Height          =   375
         Left            =   90
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   4725
         Width           =   3795
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«ŸÂ«— «·ﬂ·"
         Enabled         =   0   'False
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
         Height          =   195
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   5265
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VSFlex7Ctl.VSFlexGrid grdGroup 
         Height          =   4515
         Left            =   90
         TabIndex        =   81
         Top             =   180
         Width           =   3795
         _cx             =   6694
         _cy             =   7964
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
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
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
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
   End
   Begin VB.Frame Frame7 
      Height          =   960
      Left            =   7785
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8505
      Width           =   7305
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   10
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox xDiscount 
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
         Left            =   5130
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   9
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
         TabIndex        =   11
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "¬Ã· :"
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
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   585
         Width           =   630
      End
      Begin VB.Label xLate 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label xCash 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label17 
         Caption         =   "‰ﬁœÌ :"
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
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label14 
         Caption         =   "’«›Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   28
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label lblTotalQuant 
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label xTotalDisItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   315
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   540
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   165
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "Œ’„ ‰ﬁœÌ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "«·«Ã„«·Ì :"
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
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   21
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
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   1515
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
      Height          =   915
      Left            =   4950
      TabIndex        =   48
      Top             =   8550
      Width           =   2805
   End
End
Attribute VB_Name = "salesfrmamr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_No As String, lSave As Boolean
Dim CardTable As ADODB.Recordset, cFileHeader As String, rdPaid As New ADODB.Recordset
Public bRetvalue As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim Search31 As New Search3, oSearchItem As New Search3, bMarket As Boolean
Dim bEdit As Boolean
Dim nRoundG As Long, nRound As Long
Dim cFile As String, cFileClient, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
ItemsLookupAll Me, Search3, 10
End Sub
Private Function MyReplace(Optional nRow As Long = -1) As Boolean
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


On Error GoTo myerror
con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd nRow
con.CommitTrans
MyReplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 1)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        MakeSerial
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
        CalcTotals
    Else
        grid1.TextMatrix(grid1.Row, 10) = "1"
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Row + 1, 1
    End If
ElseIf ActiveControl.Name = CmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    xDoc_No.Text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    myUndo
    Unload Search31
'    CardTable.Find "DOC_NO = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
'    If ActiveControl.Name = CmdInform Then Search31.Hide Else Unload Search31
'    myload
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
    grid1.EditText = itemCode
    grid1.TextMatrix(grid1.Row, 1) = itemCode
    grid1.TextMatrix(grid1.Row, 2) = mySplit(itemDesca, 1, "(")
    grid1.TextMatrix(grid1.Row, 3) = "1"
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 3) = "1"
        grid1.AddItem ""
        MakeSerial
        grid1_AfterEdit grid1.Row, grid1.Col
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.TextMatrix(grid1.Row, 3) = "1"
        grid1_AfterEdit grid1.Row, grid1.Col
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
Private Sub chkBalance_Click()
addSetting "balance", chkBalance.Value, App.Path & "\other.txt"
End Sub

Private Sub chkprint_Click()
addSetting "print", chkprint.Value, App.Path & "\other.txt"
End Sub
Private Sub cmdCash_Click()
xVisa.Text = ""
xLate.Caption = ""
xCash.Caption = TurnValue(Val(xTotalDis.Caption), 0, "")
'xCash_LostFocus
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
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
MyLoad
End Sub
Private Sub CmdInform_Click()
CardLookup
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
Private Sub CmdNewInv_Click()
bAddnew = True
mydefine
On Error Resume Next
'xCode.SetFocus
grid1.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
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
myUndo
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
'doprint
On Error GoTo myerror
    If Val(xCash.Caption) = 0 And grid1.TextMatrix(1, 1) = "" Then Exit Sub
    If xDoc_No.Text = "" Then Exit Sub
    Print_Sales
    con.Execute "update file6_20h" & _
                    " SET " & _
                    " PRINT = TRUE " & _
                    " WHERE FILE6_20h.DOC_NO = " & MyParn(xDoc_No.Text)

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

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
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.CASH = " & Val(loctable!TOTAL & "") & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop
MsgBox "done..."
End Sub

Private Sub Command4_Click()
'Beep 200, 200
'console
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
    Grid1_Validate False
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
nRoundG = 2
chkprint.Value = Val(RetSetting("print", App.Path & "\other.txt"))
chkBalance.Value = Val(RetSetting("balance", App.Path & "\other.txt"))
openCon con

cFileClient = "File3_10"
If sboxSales <> "" Then
    cdefman = GetDesca("Select code from FILE6_25 WHERE BOX = " & MyParn(sboxSales))
    bEdit = True
End If


'cdefman = GetDesca("Select code from file6_25 where usercode = " & Val(nusercode))
'cDefBox = GetDesca("Select Code From file0_50 where man = " & cdefman)

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
        MyLoad
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
grid1.Col = 1
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
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim bNew As Boolean
With grid1

If Col = 1 Then GrdDesc Row

If (Col = 3 Or Col = 1) And Val(.TextMatrix(Row, 9)) > 1 And Val(.TextMatrix(Row, 3)) Mod Val(.TextMatrix(Row, 7)) = 0 Then
    grid1.TextMatrix(Row, 5) = .TextMatrix(Row, 9) * (Val(.TextMatrix(Row, 3)) / Val(.TextMatrix(Row, 7)))
ElseIf (Col = 3 Or Col = 1) Then
    grid1.TextMatrix(Row, 4) = .TextMatrix(Row, 10)
    grid1.TextMatrix(Row, 5) = Val(.TextMatrix(Row, 3)) * Val(.TextMatrix(Row, 4))
Else
    If Col <> 5 Then grid1.TextMatrix(Row, 5) = Val(.TextMatrix(Row, 3)) * Val(.TextMatrix(Row, 4))
End If
CalcTotals

If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = .Rows - 1
End If

If MyReplace(Row) Then
    HandleCntEdit
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

myloadgrd

If Row = grid1.Rows - 2 Then
    grid1.ShowCell grid1.Rows - 1, 1
    If Col = 1 Then grid1.Col = 4 Else grid1.Select .Rows - 1, 1
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    xBalance.Caption = ""
    If Not validRows(OldRow) Then grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
grid1.EditMaxLength = IIf(Col = 3, 6, 0)
End Sub
Private Sub Grid1_EnterCell()
If xPrinted.Value = 1 Or bEdit = False Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 1 Or grid1.Col = 3 Or grid1.Col = 4 Or grid1.Col = 5 Then
    grid1.Editable = flexEDKbdMouse
Else
   grid1.Editable = flexEDKbdMouse
End If
If grid1.Col = 3 And Me.chkBalance.Value = 1 Then
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
    grid1.Col = 1
End If
Grid1_EnterCell
End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
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
If KeyCode = 112 And xPrinted.Value = 0 And bEdit = True Then
    grid1.Row = grid1.Rows - 1
    grid1.Col = 1
    ItemsLookupAll Me, oSearchItem
End If
End Sub

Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    HandleCellPos KeyCode, Row, Col
End If
End Sub
Private Sub myPos()
    If grid1.Col = 3 Then
        grid1.Row = grid1.Row + 1
        grid1.Col = IIf(grid1.Row = grid1.Rows - 1, 1, 3)
   ElseIf grid1.Col = 1 Then
        grid1.Col = 3
     End If
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

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


Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Col = 3 And grid1.TextMatrix(Row, 1) <> "" And xstore.BoundText <> "" Then
'    nBalance = LastBalance(grid1.TextMatrix(Row, 1), xstore.BoundText, CON)
'    xbalanceitem.Caption = nBalance
'End If
If Col = 3 Then xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
End Sub

Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRows(grid1.Row) And grid1.Row <> grid1.Rows - 1 Then grid1.RemoveItem grid1.Row
xBalance.Caption = ""
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If grid1.EditText = "" Then
        n = Beep(1000, 1000)
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    Else
        If GetDesca("select item from file1_10 where item = " & MyParn(grid1.EditText)) = "" Then
            Dim cAlter As String
            cAlter = GetDesca("Select item from item_alter where itemsub = " & MyParn(grid1.EditText))
            If cAlter <> "" Then
                grid1.EditText = cAlter
            Else
                n = Beep(1000, 1000)
                MsgBox "ﬂÊœ «·’‰› €Ì— ”·Ì„"
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End If

If Col = 3 Or Col = 4 Then
   If Not IsNumeric(grid1.EditText) Then
        Cancel = True
        Exit Sub
   End If
End If
'If Row = grid1.Rows - 1 And Cancel = False Then
'    grid1.AddItem ""
'End If
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
End Sub
Private Function MYVALID(Optional bigMsg As Boolean = False) As Boolean
Dim I As Integer
If xDoc_No.Text = "" Then
    If Not bigMsg Then MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    If Not bigMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    If Not bigMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bigMsg Then MsgBox "·„ Ì „ «œŒ«· ﬂÊœ"
    Exit Function
End If

With grid1
'If Not validRows(, False, True) Then
    'MsgBox "«·»Ì«‰«  €Ì— ”·Ì„… «Ê ﬂ«„·…"
'    Exit Function
'End If
'DelValid , True
End With
MYVALID = True
End Function
Private Sub MyLoad(Optional bLeaveBal As Boolean = False)
On Error GoTo myerror
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xBox.BoundText = CardTable!BOX & ""
xMan.BoundText = CardTable!MAN & ""
xNotes.Text = CardTable!Notes & ""
xCode.Text = CardTable!Code & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
chkCash.Value = IIf(CardTable!cash, 1, 0)
'xusername.Text = TurnValue(CardTable!UserName, Null, "")
xDiscount.Text = TurnValue(Val(CardTable!discount & ""), 0, "")
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")
xCash.Caption = TurnValue(Val(CardTable!cash & ""), 0, "")
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
'xDate.Text = Format(IIf(Val(Format(Time, "hh")) > 4, Date, DateAdd("d", -1, Date)), "dd-mm-yyyy")
xDate.Text = Format(sysDate, "DD-MM-YYYY")
xBalance.Caption = ""
xBox.BoundText = sboxSales
xCode.Text = cDefClient
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
'xDoc_No.Tag = DefineMode
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
If Not CardTable.EOF Then MyLoad True
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
Private Sub GrdDesc(Row)
xBalance.Caption = ""
grid1.TextMatrix(Row, 7) = ""
If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Sub
Dim aret As Variant
aret = aGetDesca("select desca,price,package,unit,price2 from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 1)))
If UBound(aret) <> 0 Then
    grid1.TextMatrix(Row, 2) = aret(1) & ""
    grid1.TextMatrix(Row, 3) = 1
    grid1.TextMatrix(Row, 4) = aret(2) & ""
    grid1.TextMatrix(Row, 7) = aret(3) & ""
    grid1.TextMatrix(Row, 8) = aret(4) & ""
    grid1.TextMatrix(Row, 9) = aret(5) & ""
    grid1.TextMatrix(Row, 10) = aret(2) & ""
    xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)
End If
End Sub
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotal As Double, nDiscount As Double, nTotalitem As Double, nTotalDis As Double
With grid1
For I = 1 To grid1.Rows - 2
    If Val(.TextMatrix(I, 3)) <> 0 Then
        If Round(Val(.TextMatrix(I, 4)), nRoundG) <> Round(Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 3)), nRoundG) Then
            .TextMatrix(I, 4) = Myvalue(Round(Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 3)), nRoundG))
        End If
    Else
        .TextMatrix(I, 4) = ""
    End If
    nTotalitem = nTotalitem + Val(.TextMatrix(I, 7))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 5))
Next
'End With

'For i = 1 To grid1.Rows - 1
'    nTotalitem = nTotalitem + (Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)))
'    nDiscount = 1 - (Val(.TextMatrix(i, 6)) / 100)
'    grid1.TextMatrix(i, 5) = Myvalue(Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)) * nDiscount)
'    nTotalDisItem = nTotalDisItem + (Val(.TextMatrix(i, 3)) * Val(.TextMatrix(i, 4)) * nDiscount)
'    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 3))
'Next
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
For I = 1 To grid1.Rows - 2
    If I <> nRow And Trim(grid1.TextMatrix(I, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 0)
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
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub FixGrd()
With grid1
.FormatString = "„|" & "ﬂÊœ|" & "«·’‰‹›|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·≈Ã„«·Ì|" & "«·Œ’„|" & "«·⁄»Ê…|" & "«·ÊÕœ…|" & "”⁄— «·⁄»Ê…|" & "”⁄— «·ÊÕœ…|"
.ColWidth(0) = 500
.ColWidth(1) = 1800
.ColWidth(2) = 4500
.ColWidth(3) = 700
.ColWidth(4) = 800
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 700
.ColWidth(8) = 1100
.ColWidth(9) = 1300
.ColHidden(6) = True
.ColHidden(.Cols - 4) = True
.ColHidden(.Cols - 3) = True
.ColHidden(.Cols - 2) = True
.ColHidden(.Cols - 1) = True
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
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
    For I = nRow + 1 To grid1.Rows - 2
        con.Execute "update " & cFile & " set [row] = " & (I - 1) & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
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
'If Not xDoc_No.Enabled Then UpdateHeader
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
    MyLoad
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
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
If Not MYVALID Then Exit Sub
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
        
        temptable!str12 = !desca
        temptable!VAL1 = !TQUANT
        temptable!val2 = !PRICE
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
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert(7, 1)
With grid1
    For I = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(I, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(I, 3))

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(I, 4))

        aInsert(4, 0) = "TOTAL"
        aInsert(4, 1) = Val(.TextMatrix(I, 5))

        aInsert(5, 0) = "Discount"
        aInsert(5, 1) = Val(.TextMatrix(I, 6))

        aInsert(6, 0) = "PACKAGE"
        aInsert(6, 1) = Val(.TextMatrix(I, 7))

        aInsert(7, 0) = "Cost"
        aInsert(7, 1) = LastCostDate(grid1.TextMatrix(I, 1), xDate.Text, con)
        
       
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
    cField1 = "case when " & cFile & ".Discount = 0 then Null else " & cFile & ".Discount end "
    cString = "SELECT FILE6_20.ITEM, FILE1_10.DESCA, FILE6_20.Quant,FILE6_20.Price,FILE6_20.TOTAL," & cField1 & ",FILE6_20.PACKAGE,FILE1_10.[UNIT],FILE1_10.PRICE2,FILE1_10.PRICE, ID " & _
              " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
    cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.Text)
    DATA3.RecordSource = cString
    DATA3.Refresh
    grid1.AddItem ""
    MakeSerial
End With
Handlecontrols LoadMode
CalcTotals
FixGrd
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
If Not MYVALID Then Exit Function
CalcTotals
If Not MyReplace Then Exit Function
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
'    CardTable.Requery
'    CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
'    If CardTable.EOF Then CardTable.MoveLast
'    myload
    myloadgrd
    xDoc_No.Tag = LoadMode
    xDoc_No.Enabled = False
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
    grdGroup.TextMatrix(grdGroup.Rows - 1, 0) = loctable!desca & ""
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
For I = 0 To .Rows - 1
    cString = "UPDATE FILE1_50 SET FILE1_50.SERIAL = " & I & _
               " WHERE CODE = " & .TextMatrix(I, .Cols - 1)
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
Private Function olddoprint() As Boolean
On Error GoTo myerror
Dim aHeader(2)
If Not MYVALID Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1
For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!Date1 = DateFix(xDate.Text)
    temptable!str5 = TurnValue(xtime.Caption)
    
    If xNotes.Text <> "" Then temptable!str2 = "«·”«œ… : " & xNotes.Text
    temptable!Str11 = "›«œÌ „«—ﬂ  "
    
    temptable!str12 = TurnValue(turn(xSuper.Text, "«·„‘—› : ") & xSuper.Text)
    temptable!Str13 = TurnValue(turn(xCarNo.Text, "—ﬁ„ «·”Ì«—… : ") & xCarNo.Text)
    temptable!Str14 = TurnValue(turn(xcardesca.Text, "„ÊœÌ· : ") & xcardesca.Text)
    
    temptable!str3 = ArbString(Val(xDoc_No.Text))
    temptable!str6 = .TextMatrix(I, 2)
    temptable!str4 = TurnValue(xMan.Text)
    temptable!VAL1 = Val(.TextMatrix(I, 3))
    temptable!val2 = Val(.TextMatrix(I, 4))
    temptable!val3 = Val(.TextMatrix(I, 5))
    temptable!Str11 = ArbString("„Ê»Ì· : 01003447035")
    temptable!val4 = Val(xTotalItem.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xCash.Caption)
    temptable!Val7 = Val(xPay.Caption)
    temptable!Val8 = Val(xRest.Caption)
    temptable.Update
Next I
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
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    HandleCellPos KeyCode, Row, Col
End If
End Sub
Private Sub HandleCellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
If (Col = 1) Then
    KeyCode = 0
    If IsNumeric(grid1.TextMatrix(Row, 1)) Then grid1.Col = 10
ElseIf Col = 10 Or Col = 11 Then
    If Row = grid1.Rows - 2 Then
        KeyCode = 0
        grid1.Select grid1.Rows - 1, 1
    End If
End If
End Sub
Private Function validRow(Row As Long, Optional bigMsg As Boolean = False, Optional bIgMsgsub As Boolean = False) As Boolean
With grid1
If Not MYVALID(bigMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
End With
validRow = True
End Function

Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
cmdSave.Enabled = (bEdit) And xPrinted.Value = 0 And grid1.Rows > 2
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT FILE6_20H.*,FILE3_10.DESCA AS CLIENTDESCA FROM FILE6_20H LEFT JOIN FILE3_10 ON FILE6_20H.Code = FILE3_10.CODE"
If sDoc_No <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_No)
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
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
Private Sub Print_Sales()
On Error GoTo myerror
Dim TargetTable As New ADODB.Recordset
Dim nTPrice As Double
contemp.Execute "DELETE FROM TEMP"
TargetTable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTableDirect
nTPrice = 0

With grid1
For I = 1 To .Rows - 2
    TargetTable.AddNew
   
    TargetTable!Date1 = xDate.Text
    TargetTable!date2 = xtime.Text
    If xNotes.Text <> "" Then TargetTable!str2 = "«·”«œ… : " & xNotes.Text
    TargetTable!str3 = Val(xDoc_No.Text)
    TargetTable!str6 = .TextMatrix(I, 2)
    TargetTable!str4 = TurnValue(xusername.Text)
    TargetTable!VAL1 = Val(.TextMatrix(I, 3))
    TargetTable!val2 = Val(.TextMatrix(I, 4))
    TargetTable!val3 = Val(.TextMatrix(I, 5))
    
    TargetTable!val4 = xTotalItem.Caption
    TargetTable!val5 = Val(xDiscount.Text)
    TargetTable!Val6 = Val(xCash.Caption)
    TargetTable!Val7 = Val(xPay.Text)
    TargetTable!Val8 = Val(xChang.Text)
    TargetTable.Update
Next I
End With
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
REPORT1.WindowShowPrintSetupBtn = True
REPORT1.PrinterName = pDevice
REPORT1.ReportFileName = App.Path & "\reports\SALES_BON.rpt"
main.REPORT1.DataFiles(0) = tempPath
REPORT1.Destination = crptToPrinter
REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

