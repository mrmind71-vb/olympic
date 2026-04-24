VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Vs_Inv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì— «·„»Ì⁄« "
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNewinv 
      BackColor       =   &H00C0FFFF&
      Caption         =   "›« Ê—… ÃœÌœ…"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6885
      Width           =   1860
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Õ›Ÿ «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   6885
      Width           =   1995
   End
   Begin VB.Frame Frame6 
      Height          =   960
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5895
      Width           =   5775
      Begin VB.TextBox xTotItem 
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
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1230
      End
      Begin VB.TextBox xCash 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   135
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   135
         Width           =   1230
      End
      Begin VB.TextBox xVisa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   330
         Left            =   135
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   495
         Width           =   1230
      End
      Begin VB.TextBox xDisc 
         Alignment       =   2  'Center
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
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   270
         Width           =   1620
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’‹‹‹‹‹‹‹‹‹‹‹‹„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ›Ì‹‹‹“«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   585
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ‰ﬁ‹‹‹œÏ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      Height          =   1725
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   585
      Width           =   7575
      Begin VB.TextBox xTime 
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
         Left            =   180
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox xDoc_No 
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
         Left            =   4860
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1500
      End
      Begin VB.TextBox xDate 
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
         Left            =   1170
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xCode 
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
         Left            =   4860
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox xDoc 
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
         Left            =   4860
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1260
         Width           =   1500
      End
      Begin VB.TextBox xLevel 
         Alignment       =   2  'Center
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
         ForeColor       =   &H002C53C9&
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1260
         Width           =   870
      End
      Begin VB.TextBox xMan 
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
         Left            =   4860
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   555
      End
      Begin VB.Label xClientBalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label xclientDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   540
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6495
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ‹‹‹‹‹‹‹‹Êœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6495
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ì‹‹‹‹‹‹‹œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·»Ì«‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6495
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1305
         Width           =   765
      End
      Begin VB.Label LMAN 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»«∆⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6495
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   945
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label xManDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   900
         Visible         =   0   'False
         Width           =   4650
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1140
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1170
      Width           =   1635
      Begin VB.CommandButton Cmd_Cust 
         Caption         =   "»Ì«‰«  ⁄„·«¡"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   630
         Width           =   1545
      End
      Begin VB.CommandButton Cmd_Item 
         BackColor       =   &H00C0FFFF&
         Caption         =   "»Ì«‰«  «·√’‰«›"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.TextBox xUser 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   2670
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·›« Ê—… "
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1305
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   4830
      Begin VB.CommandButton CmdAddPrint 
         BackColor       =   &H00CEC2AA&
         Caption         =   " —ÕÌ· »«—ﬂÊœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   135
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÿ»«⁄… ≈–‰  ”·Ì„"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   135
         Width           =   1680
      End
      Begin VB.CommandButton Cmd_Print 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÿ»«⁄… «·›« Ê—…"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   135
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   390
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   9000
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox xCustName 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E8&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   945
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8190
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIX"
      Height          =   240
      Left            =   300
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   8415
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5895
      Width           =   6750
      Begin VB.TextBox xPrice 
         Alignment       =   2  'Center
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox xQty 
         Alignment       =   2  'Center
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox xItem 
         Alignment       =   2  'Center
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1455
      End
      Begin MSDBCtls.DBCombo xStore 
         Bindings        =   "Vs_invNew.frx":0000
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Top             =   270
         Width           =   2985
         _ExtentX        =   5265
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·»Ì⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   750
      End
      Begin VB.Label xBalItemInv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         TabIndex        =   15
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label xDescItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   690
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   4245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂ„Ì‹‹‹…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   990
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·’‰‹‹‹‹›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   555
      Left            =   540
      TabIndex        =   19
      Top             =   1485
      Visible         =   0   'False
      Width           =   825
      _cx             =   1455
      _cy             =   979
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   9
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
      MarginLeft      =   100
      MarginTop       =   100
      MarginRight     =   100
      MarginBottom    =   100
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
      Zoom            =   -1.51380231522707
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   200
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   1
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid ItemInv 
      Height          =   3525
      Left            =   315
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2340
      Width           =   14865
      _cx             =   26220
      _cy             =   6218
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
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
   Begin VB.Frame Frame8 
      Height          =   570
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   6795
      Width           =   1890
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
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "«Ê·"
         Top             =   135
         Visible         =   0   'False
         Width           =   435
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
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   135
         Visible         =   0   'False
         Width           =   435
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
         Height          =   360
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
         Top             =   135
         Visible         =   0   'False
         Width           =   435
      End
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
         Height          =   360
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
         Top             =   135
         Width           =   435
      End
   End
   Begin VB.Label xQ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   315
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   135
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "Vs_Inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UnInvTable As Recordset
Dim DocTable As Recordset, ClientTable As Recordset, itemTable As Recordset
Dim SalesTable As Recordset
Dim CustTable As Recordset

Dim PaySalTable As Recordset
Dim MoveItemTable As Recordset
Dim cStore As String, SerFileName As String
Dim BalClient As Recordset
Dim ItemCode As Recordset

Dim GrTable As Recordset
Dim FlagTable As Recordset
Dim storeTable As Recordset, COUNTINVTOTAL As Double
Dim formMode, cSerItem As String
Dim itemMoveType As String, ClientMoveType As String
Dim myFileName As String, nDiscount As Byte, sStore As String, nPrice As Double
Dim cCost As String
Dim cStrStore As String
Const NewInvMode = 4, applyMode = 5
Sub DocValid()
If xDoc_No.Text = "" Then Exit Sub
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.INDEX = "nDocItem"

DocTable.Seek ">=", xDoc_No.Text
If DocTable.NoMatch Then Exit Sub
If DocTable.doc_no <> xDoc_No Then Exit Sub

xCash.Enabled = False
xDisc.Enabled = False
xVisa.Enabled = False

xDoc_No.Enabled = True
xCode.Enabled = True
xMan.Enabled = True

xDate.Enabled = True
xDoc.Enabled = True

xItem.Enabled = True
xPrice.Enabled = True
xQty.Enabled = True

ApplyProc
End Sub
Sub editProc()
formMode = Editmode
End Sub
Sub EmptyProc()
formMode = EmptyMode
xCode.Text = ""
'xMan.Text = ""
xDoc_No.Text = ""
'xRemark.Text = ""
xDoc.Text = ""
xDate.Text = Format(Date, "dd-mm-yyyy")
xLevel.Text = ""
ItemInv.Rows = 1
End Sub
Sub AddProc()
    formMode = addmode
    ItemInv.AddItem ""
End Sub
Sub Fillgrd()
Dim nT_Item, nT_DOZ, nT_Disc, nT_Inv As Double
nT_Item = 0
nT_DOZ = 0
nT_Inv = 0
ItemInv.Rows = 1
i = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow

xTotItem.Text = ""
xDisc.Text = ""
nT_Item = 0

DocTable.Seek ">=", xDoc_No.Text
If DocTable.NoMatch Then Exit Sub
If DocTable.doc_no <> xDoc_No.Text Then Exit Sub
xCustName.Text = ""
If publicFlag = "1" Then
    cCust = TurnValue(DocTable.Cust, Null, "")
    CustTable.FindFirst " CODE = " & MyParn(cCust)
    If Not CustTable.NoMatch Then xCustName.Text = CustTable.desca
End If
Do While True
    nT_Inv = nT_Inv + TurnValue(DocTable.total, Null, 0)
    If DocTable.Store <> "zz" Then
        If xStore.BoundText <> "" Then
            xStore.BoundText = TurnValue(DocTable.Store, Null, "")
        End If
        
        .AddItem ""
        .TextMatrix(i, 0) = TurnValue(DocTable.Store, Null, "")
        .TextMatrix(i, 1) = TurnValue(DocTable.Item, Null, "")
        itemTable.FindFirst " item = " & MyParn(.TextMatrix(i, 1))
        If Not itemTable.NoMatch Then .TextMatrix(i, 2) = itemTable.desca
        .TextMatrix(i, 3) = Format(DocTable.Quant, "#0")
        .TextMatrix(i, 5) = TurnValue(Format(DocTable.price, "##0.00"), Null, "")
        .TextMatrix(i, 6) = TurnValue(Format(DocTable.total, "##0.00"), Null, "")
        nT_Item = nT_Item + DocTable.total
        nT_DOZ = nT_DOZ + Val(.TextMatrix(i, 3))
        xTime.Text = Format(DocTable.Time, "SHORT TIME")
        xUser.Text = TurnValue(DocTable.User, Null, "")
        i = i + 1
    Else
        nT_Disc = nT_Disc + DocTable.total
    End If
    DocTable.MoveNext
    If DocTable.EOF Then Exit Do
    If DocTable.doc_no <> xDoc_No.Text Then Exit Do
Loop
nCrow = 0
End With
xTotItem.Text = Format(nT_Item, "##0.00")
xDisc.Text = Format(nT_Disc * -1, "##0.00")
If publicFlag = 1 Then
    PaySalTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not PaySalTable.NoMatch Then
        xCash.Text = Format(TurnValue(PaySalTable.CASH, Null, 0), "#0.00")
        xVisa.Text = Format(TurnValue(PaySalTable.VISA, Null, 0), "#0.00")
    Else
        xCash.Text = Val(xTotItem.Text) - Val(xDisc.Text)
    End If
Else
    xCash.Text = Val(xTotItem.Text) - Val(xDisc.Text)
End If
xQ.Caption = nT_DOZ
End Sub
Sub ItemsLookup()
'    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(4)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] ,FACTCODE AS [ﬂÊœ «·„’‰⁄] , price as [„” Â·ﬂ] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')   OR  ITEM Like('*cFilter*')   OR  factcode Like('*cFilter*')    "
    Generalarray(4) = "Order by Item"
    GrdArray(1) = 1000
    GrdArray(2) = 3500
    GrdArray(3) = 1500
    GrdArray(4) = 1000
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Function MyReplace()
Dim LPR As Boolean
LPR = False
MyReplace = True
If publicFlag = "1" Then
    MsgBox (" ﬁÌ„… «·„” Õﬁ ··œ›⁄ " & Format(Val(xTotItem.Text) - Val((xDisc.Text)), "#0.00"))
    UnInvTable.Seek "=", xDoc_No.Text
    If Not UnInvTable.NoMatch Then
        If MsgBox(" «·»Ê‰ „⁄·ﬁ ”Ê› Ì „ Õ›Ÿ «·»Ê‰", vbOKCancel) = vbOK Then UnInvTable.Delete
    End If
Else
    MsgBox "  „ Õ›Ÿ «·›« Ê—… "
End If

If publicFlag = "1" Then
    If ClientTable.CASH Then
        If Str(Val(xCash.Text) + Val(xVisa.Text)) <> Str(Val(xTotItem.Text) - Val(xDisc.Text)) Then
            xCash.Text = Val(xTotItem.Text) - Val(xDisc.Text) - Val(xVisa.Text)
        End If
        
        PaySalTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
        If Not PaySalTable.NoMatch Then
            PaySalTable.Edit
        Else
            PaySalTable.AddNew
        End If
        PaySalTable.doc_no = xDoc_No.Text
        PaySalTable.CASH = Val(xCash.Text)
        PaySalTable.VISA = Val(xVisa.Text)
        PaySalTable.total = Val(xTotItem.Text) - Val(xDisc.Text)
        PaySalTable.Update
    Else
        PaySalTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
        If Not PaySalTable.NoMatch Then PaySalTable.Delete
    End If
    mydb.Execute "delete * from file6_20 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz'"
ElseIf publicFlag = "3" Then
     mydb.Execute "delete * from file7_20 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz'"
End If

Select Case publicFlag
    Case "1"
        mydb.Execute " delete * from file6_20 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz' "
        mydb.Execute " UPDATE FILE6_20 SET FILE6_20.[DATE] = " & DateSql(xDate.Text) & " , FILE6_20.CODE = " & MyParn(xCode.Text) & " , FILE6_20.MAN = " & MyParn(xMan.Text) & ", doc = " & addstring(xDoc.Text) & " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
    Case "2"
        mydb.Execute " delete * from file6_10 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz' "
        mydb.Execute " UPDATE FILE6_10 SET FILE6_10.[DATE] = " & DateSql(xDate.Text) & " , FILE6_10.CODE = " & MyParn(xCode.Text) & ", doc = " & addstring(xDoc.Text) & " WHERE FILE6_10.DOC_NO = " & MyParn(xDoc_No.Text)
    Case "3"
         mydb.Execute "delete * from file7_20 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz'"
        mydb.Execute " UPDATE FILE7_20 SET FILE7_20.[DATE] = " & DateSql(xDate.Text) & " , FILE7_20.CODE = " & MyParn(xCode.Text) & ", doc = " & addstring(xDoc.Text) & " WHERE FILE7_20.DOC_NO = " & MyParn(xDoc_No.Text)
    Case "4"
         mydb.Execute "delete * from file6_11 where Doc_No = " & MyParn(xDoc_No.Text) & " and store = 'zz'"
        mydb.Execute " UPDATE FILE6_11 SET FILE6_11.[DATE] = " & DateSql(xDate.Text) & " , FILE6_11.CODE = " & MyParn(xCode.Text) & ", doc = " & addstring(xDoc.Text) & " WHERE FILE6_11.DOC_NO = " & MyParn(xDoc_No.Text)
End Select

If Val(xDisc.Text) <> 0 Then
    Select Case publicFlag
        Case "1"
            cString = "Insert Into FILE6_20 (Doc_no,[Date],Code,MAN,Store,Total) " & _
                            " Values(" & _
                            addstring(xDoc_No.Text) & "," & _
                            DateSql(xDate.Text) & "," & _
                            addstring(xCode.Text) & "," & _
                            addstring(xMan.Text) & "," & _
                            addstring("zz") & "," & _
                            addvalue(Val(xDisc.Text) * -1) & _
                            ")"
            mydb.Execute cString
        Case "2"
            cString = "Insert Into FILE6_10 (Doc_no,[Date],Code,Store,Total) " & _
                            " Values(" & _
                            addstring(xDoc_No.Text) & "," & _
                            DateSql(xDate.Text) & "," & _
                            addstring(xCode.Text) & "," & _
                            addstring("zz") & "," & _
                            addvalue(Val(xDisc.Text) * -1) & _
                            ")"
            mydb.Execute cString
        Case "3"
            cString = "Insert Into FILE7_20 (Doc_no,[Date],Code,Store,Total) " & _
                            " Values(" & _
                            addstring(xDoc_No.Text) & "," & _
                            DateSql(xDate.Text) & "," & _
                            addstring(xCode.Text) & "," & _
                            addstring("zz") & "," & _
                            addvalue(Val(xDisc.Text) * -1) & _
                            ")"
            mydb.Execute cString
        Case "4"
            cString = "Insert Into FILE6_11 (Doc_no,[Date],Code,Store,Total) " & _
                            " Values(" & _
                            addstring(xDoc_No.Text) & "," & _
                            DateSql(xDate.Text) & "," & _
                            addstring(xCode.Text) & "," & _
                            addstring("zz") & "," & _
                            addvalue(Val(xDisc.Text) * -1) & _
                            ")"
            mydb.Execute cString
    End Select
End If
    
    
    ' ≈‰‘«¡ Õ—ﬂ… √’‰«› «·›« Ê—…
'   If lPr And Not lManger Then Cmd_Print_Click
End Function
Function MyReplace2()
If publicFlag = "1" Then
    mydb.Execute " UPDATE FILE6_20 SET FILE6_20.[DATE] = " & DateSql(xDate.Text) & " , FILE6_20.CODE = " & MyParn(xCode.Text) & " , FILE6_20.MAN = " & MyParn(xMan.Text) & ", doc = " & addstring(xDoc.Text) & " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
    If ClientTable.CASH Then
        If Str(Val(xCash.Text) + Val(xVisa.Text)) <> Str(Val(xTotItem.Text) - Val(xDisc.Text)) Then
            xCash.Text = Val(xTotItem.Text) - Val(xDisc.Text) - Val(xVisa.Text)
        End If
        
        PaySalTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
        If Not PaySalTable.NoMatch Then
            PaySalTable.Edit
        Else
            PaySalTable.AddNew
        End If
        PaySalTable.doc_no = xDoc_No.Text
        PaySalTable.CASH = Val(xCash.Text)
        PaySalTable.VISA = Val(xVisa.Text)
        PaySalTable.total = Val(xTotItem.Text) - Val(xDisc.Text)
        PaySalTable.Update
    Else
        PaySalTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
        If Not PaySalTable.NoMatch Then PaySalTable.Delete
    End If
End If
End Function
Sub ApplyProc()
'xCash.Enabled = False
xDisc.Enabled = False
'xVisa.Enabled = False
'xDoc_No.Enabled = True
'xCode.Enabled = True
'xDate.Enabled = True
'xDoc.Enabled = True
DocTable.Seek ">=", xDoc_No.Text
If DocTable.NoMatch Then Exit Sub
If DocTable.doc_no <> xDoc_No.Text Then Exit Sub
If Not DocTable.EOF Then
If DocTable.NoMatch Then
    EmptyProc
Else
    ViewClose
    xCode.Text = DocTable.CODE
    
    xMan.Text = TurnValue(DocTable.MAN, Null, "")
    xManDescA.Caption = SayCode(FlagTable, 6, xMan.Text)

    xDoc.Text = TurnValue(DocTable.doc, Null, "")
    If DocTable.Store <> "zz" Then
        xStore.BoundText = TurnValue(DocTable.Store, Null, "")
    End If
      
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    ClientTable.FindFirst "Code = " & MyParn(xCode.Text)
    If Not ClientTable.NoMatch Then
        xclientDescA.Caption = ClientTable.desca
         
        If publicFlag = "1" Then
            xLevel.Text = TurnValue(ClientTable.price, Null, 0)
        Else
            xLevel.Text = "1"
        End If
    End If
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.grid1, 0)
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·›« Ê—…"
    MYVALID = False
End If
If xDate.Text = "" Or Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
If xCode.Text = "" Then
    MsgBox " ”ÃÌ· «·≈”„"
    MYVALID = False
End If

ClientTable.FindFirst "code = " & MyParn(xCode.Text)
If ClientTable.NoMatch Then
    MsgBox "ﬂÊœ €Ì— „”Ã·"
    MYVALID = False
End If
If publicFlag = "1" And ClientTable.CASH Then
    If Str(Val(xVisa.Text) + Val(xCash.Text) + Val(xDisc.Text)) <> Str(Val(xTotItem.Text)) Then
        MsgBox "—«Ã⁄ ﬁÌ„… «·»Ê‰ „⁄ ﬁÌ„… «·„”œœ ‰ﬁœ« Ê ›Ì“«"
        Exit Function
    End If
End If
End Function
Sub Undoinv()
Select Case formMode
Case addmode
    InvGrid.Rows = InvGrid.Rows - 1
    dispProc
Case Editmode
    dispProc
Case EmptyMode
    
End Select
End Sub
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 2
End Sub
Private Sub CalcDisc_Click()
    With ItemInv
    Dim nRateDisc    As Double
    Dim nDisc    As Double
    nDisc = 0
    nRateDisc = 0
    nRateDisc = Val(xRateDisc.Text)
    If nRateDisc <> 0 Then
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "SS" Then
                nDisc = nDisc + Val(.TextMatrix(i, 9))
            End If
        Next i
    End If
    End With
    
    xT_Inv.Caption = Format(Val(xT_Item.Caption) - Val(xDisc.Text), "##0.00")
End Sub

Private Sub Cmd_Cust_Click()
    ShowCust.Show 1
End Sub
Private Sub CMD_ITEM_Click()
    items.Show 1
End Sub
Private Sub Cmd_Print_Click()
Dim TargetTable As Recordset
Dim nTQ As Double
nTQ = 0
tempdb.Execute "DELETE * FROM TEMP"
With ItemInv
Set TargetTable = tempdb.OpenRecordset("TEMP")
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) <> "" Then
            TargetTable.AddNew
            TargetTable.str1 = Chr254(xDoc_No.Text)
            TargetTable.str3 = xclientDescA.Caption
            TargetTable.str2 = Chr254(xCode.Text)
            TargetTable.Date1 = xDate.Text
            itemTable.FindFirst " item = " & MyParn(.TextMatrix(i, 1))
            If Not itemTable.NoMatch Then
                TargetTable.str7 = Chr254(Format(itemTable.price, "##0.00"))
                
                TargetTable.STR5 = .TextMatrix(i, 1)
                TargetTable.str9 = Chr(254) & itemTable.desca & Chr(254)
                TargetTable.str4 = itemTable.FACTCODE
                TargetTable.str6 = itemTable.Group
                
                GrTable.FindFirst " CODE = " & MyParn(itemTable.Group)
                TargetTable.str8 = GrTable.desca
                TargetTable.Str11 = Chr254(Format(.TextMatrix(i, 3), "##0"))
                
                
                TargetTable.str12 = Chr254(Format(.TextMatrix(i, 4), "##0"))
                TargetTable.STR13 = Chr254(Format(.TextMatrix(i, 5), "##0.00"))
                TargetTable.STR14 = Chr254(Format(.TextMatrix(i, 6), "##0.00"))
                
                TargetTable.str16 = Chr254(Format(itemTable.price, "##0.00"))
                TargetTable.STR17 = Chr254(Format(Me.xTotItem.Text, "##0.00"))
                TargetTable.STR18 = Chr254(Format(xQ.Caption, "##0.00"))
                If Val(xDisc.Text) <> 0 Then
                    TargetTable.str12 = "Œ’‹‹‹„"
                    TargetTable.STR15 = Chr254(Format(xDisc.Text, "#0.00"))
                    TargetTable.str10 = "’«›Ï «·›« Ê—…"
                    TargetTable.STR19 = Chr254(Format(Val(xTotItem) - Val(xDisc.Text), "##0.00"))
                End If
            End If
            TargetTable.Update
        End If
    Next i

End With
myws.BeginTrans
myws.CommitTrans
REPORT1.WindowState = crptMaximized
If publicFlag = 1 Then
    REPORT1.ReportFileName = PublicPath & "\Reports\PRINTSAL.rpt"
ElseIf publicFlag = 3 Then
    REPORT1.ReportFileName = PublicPath & "\Reports\PRINTPURCH.rpt"
ElseIf publicFlag = 4 Then
    REPORT1.ReportFileName = PublicPath & "\Reports\RPURCH.rpt"
End If
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1

End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·›« Ê—… »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        EmptyProc
    End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
'On Error Resume Next

CmdDelInv.Enabled = True
CmdSave.Enabled = True
ItemInv.Rows = 1
ItemInv.Rows = 2
xCode.Text = ""
If publicFlag = "1" Then xCode.Text = "000000"
xClientBalance.Caption = ""
xclientDescA.Caption = ""
xDoc_No.Enabled = True
xStore.Enabled = True
storeTable.MoveFirst
xStore.BoundText = storeTable.CODE

'xCash.Enabled = False
xDisc.Enabled = False
'xVisa.Enabled = False

xDoc_No.Enabled = True
xCode.Enabled = True
xMan.Enabled = True

xDate.Enabled = True
xDoc.Enabled = True

xItem.Enabled = True
xPrice.Enabled = True
xQty.Enabled = True

Me.xCash.Text = ""
Me.xDisc.Text = ""
Me.xTotItem.Text = ""
xDoc.Text = ""
'xMan.Text = ""
If Hour(Time) < 4 Then
    xDate.Text = Format(DateAdd("D", -1, Date), "dd-mm-yyyy")
Else
    xDate.Text = Format(Date, "dd-mm-yyyy")
End If
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
xTotItem.Text = ""
xDisc.Text = 0
On Error Resume Next
xDoc_No.SetFocus
Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
If publicFlag <> 1 Then Exit Sub

'If Val(xTotItem.Text) >= 100 Then
'    Load Cust
'
'    Cust.Show 1
'    CustTable.Requery
'    If cCust <> "" Then
'        mydb.Execute " UPDATE FILE6_20 SET FILE6_20.[CUST] = " & MyParn(cCust) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
'    End If
'
'End If
UnInvTable.Seek "=", xDoc_No.Text
If UnInvTable.NoMatch Then
    If publicFlag = 1 And ClientTable.CASH Then
       If nCountPrint > 0 Then
            For i = 1 To nCountPrint
                myprint
            Next i
        End If
    End If
End If
'Fillgrd
'CmdNewInv_Click
End Sub
Private Sub Command1_Click()
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
With ItemInv
Set TargetTable = tempdb.OpenRecordset("TEMP")
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) <> "" Then
            TargetTable.AddNew
            TargetTable.str1 = Chr254(xDoc_No.Text)
            TargetTable.str3 = xclientDescA.Caption
            TargetTable.str2 = Chr254(xCode.Text)
            TargetTable.Date1 = xDate.Text
            itemTable.FindFirst " item = " & MyParn(.TextMatrix(i, 1))
            If Not itemTable.NoMatch Then
                TargetTable.STR5 = .TextMatrix(i, 1)
                TargetTable.str6 = itemTable.Group
                GrTable.FindFirst " CODE = " & MyParn(itemTable.Group)
                TargetTable.str8 = GrTable.desca
                TargetTable.str9 = Chr(254) & .TextMatrix(i, 2) & Chr(254)
                TargetTable.str4 = itemTable.FACTCODE
                TargetTable.Str11 = Chr254(Format(.TextMatrix(i, 3), "##0"))
                
                TargetTable.STR13 = Chr254(Format(.TextMatrix(i, 5), "##0.00"))
                TargetTable.STR14 = Chr254(Format(.TextMatrix(i, 6), "##0.00"))
                
                TargetTable.str16 = Chr254(Format(itemTable.price, "##0.00"))
            
            End If
            TargetTable.Update
        End If
    Next i

End With
myws.BeginTrans
myws.CommitTrans
REPORT1.WindowState = crptMaximized
REPORT1.ReportFileName = PublicPath & "\Reports\PRINTSAL3.rpt"
REPORT1.DataFiles(0) = cPathTemp
REPORT1.Action = 1
End Sub
Private Sub Command2_Click()
'Exit Sub
Dim dNewDate As Date
Dim CdOC As String
For i = 1 To 5452
    CdOC = RetZero(i, 6)
    Me.Caption = CdOC
    xDoc_No.Text = CdOC
    DocTable.INDEX = "nDocItem"
    DocTable.Seek ">=", CdOC
    
'    If Not DocTable.NoMatch Then
'        If DocTable!Date > DateValue("1-1-2013") Then
'            xDate.Text = DocTable!Date
'            dNewDate = Mid(xDate.Text, 9, 2) & "-" & Mid(xDate.Text, 4, 2) & "-20" & Mid(xDate.Text, 1, 2)
'            mydb.Execute " UPDATE FILE6_20 SET FILE6_20.[DATE] = " & DateSql(dNewDate) & " WHERE FILE6_20.DOC_NO = " & MyParn(xDoc_No.Text)
'        End If
'    End If
        
    If Not DocTable.NoMatch Then
        If Trim(DocTable!doc_no) <> Trim(CdOC) Then
            PaySalTable.FindFirst " doc_no = " & MyParn(CdOC)
            If Not PaySalTable.NoMatch Then PaySalTable.Delete
        Else
            DocValid
            MyReplace2
        End If
    End If
Next i



End Sub
Private Sub Command3_Click()
Dim PurchTable As Recordset
Set PurchTable = mydb.OpenRecordset("FILE7_20")

Dim MyItem As Recordset
Set MyItem = mydb.OpenRecordset("FILE1_10")
MyItem.INDEX = "nItem"

PurchTable.INDEX = "nItem"
With PurchTable
    Do While Not .EOF
        MyItem.Seek "=", .Item
        If Not MyItem.NoMatch Then
            MyItem.Edit
            MyItem.COST = .price
            MyItem.Update
        End If
        .MoveNext
    Loop
End With

With DocTable
    .MoveFirst
    Do While Not DocTable.EOF
        Me.Caption = DocTable.doc_no
        nCost = 0
        MyItem.Seek "=", .Item
        If Not MyItem.NoMatch Then nCost = MyItem.COST
        
        PurchTable.Seek ">=", DocTable.Item, DocTable.Date
        If Not PurchTable.NoMatch Then
            PurchTable.MovePrevious
            If PurchTable.Item = .Item Then
                nCost = PurchTable!price
            End If
        End If
        .Edit
        .COST = nCost
        .Update
        .MoveNext
    Loop
End With
End Sub

Private Sub Command4_Click()
With DocTable
    .MoveFirst
    Do While Not DocTable.EOF
        Me.Caption = DocTable.doc_no
        .Edit
        .doc_no = RetZero(.doc_no, 6)
        .Update
        .MoveNext
    Loop
End With

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Cmd_Print.Visible = bopt3
Command1.Visible = bopt3
cmd_item.Visible = Main.XBDATA.Visible
Select Case publicFlag
    Case 1 ' „»Ì⁄« 
        LMAN.Visible = True
        xMan.Visible = True
        xManDescA.Visible = True
        Set UnInvTable = mydb.OpenRecordset("FILE6_00")
        UnInvTable.INDEX = "nDoc"
        itemMoveType = "6"
        myFileName = "File6_20"
        ClientMoveType = "4"
        Set ClientTable = mydb.OpenRecordset("File3_10", dbOpenDynaset)
        Set CustTable = mydb.OpenRecordset("File3_20", dbOpenDynaset)
        Set ItemCode = mydb.OpenRecordset("ITEMCODE")
        Set PaySalTable = mydb.OpenRecordset("SELECT * FROM File6_22 ORDER BY DOC_NO ", dbOpenDynaset)
        Me.Caption = "›« Ê—… „»Ì⁄« "
        CmdAddPrint.Caption = "Ã. „»Ì⁄«  ÌÊ„Ì…"
    Case 2 ' „—œÊœ „»Ì⁄« 
        itemMoveType = "3"
        myFileName = "File6_10"
        ClientMoveType = "5"
        Set SalesTable = mydb.OpenRecordset("FILE6_20")
        SalesTable.INDEX = "nCodeItem"
        Set ClientTable = mydb.OpenRecordset("File3_10", dbOpenDynaset)
        Me.Caption = "›« Ê—… „—œÊœ „»Ì⁄« "
        CmdAddPrint.Visible = False
    Case 3 ' „‘ —Ì« 
        itemMoveType = "2"
        myFileName = "File7_20"
        ClientMoveType = "4"
        Set ClientTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
        Me.Caption = "›« Ê—… „‘ —Ì« "
    Case 4
        itemMoveType = "7"
        myFileName = "File6_11"
        ClientMoveType = "5"
        Set ClientTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
        Me.Caption = "›« Ê—… „—œÊœ «·„‘ —Ì«  "
        CmdAddPrint.Visible = False
End Select
Set FlagTable = mydb.OpenRecordset("SELECT * FROM File1_70 ", dbOpenDynaset)


'CmdInform.Visible = bopt1
cmdPrevious.Visible = bopt1
cmdFirst.Visible = bopt1
cmdNext.Visible = bopt1

xDoc_No.Locked = Not bopt1

Command1.Visible = bopt1

data1.DatabaseName = MdbPath
data1.RecordSource = "SELECT * FROM FILE0_50"
xStore.BoundColumn = "Code"
xStore.ListField = "DescA"

Set DocTable = mydb.OpenRecordset(myFileName)
Set GrTable = mydb.OpenRecordset("SELECT * FROM FILE1_50")
DocTable.INDEX = "nDocItem"

Set storeTable = mydb.OpenRecordset("Stores", dbOpenDynaset)
Set itemTable = mydb.OpenRecordset("file1_10", dbOpenDynaset)
Set ChargesTable = mydb.CreateSnapshot("Select * From File8_70 ", dbOpenDynaset)
Set MoveItemTable = mydb.OpenRecordset("FILE1_11")
xDate.Text = Format(Date, "dd-mm-yyyy")

storeTable.MoveFirst
xStore.BoundText = storeTable.CODE

cStrStore = StrStore
If DocTable.RecordCount > 0 Then
    DocTable.MoveFirst
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 7
    .Rows = 1
    .TextMatrix(0, 0) = "„Œ“‰"
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·’‰‹‹‹‹›"
    .TextMatrix(0, 3) = "«·ﬂ„Ì…"
    .TextMatrix(0, 5) = "«·”⁄—"
    .TextMatrix(0, 6) = "«·≈Ã„«·Ï"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1000
    .ColWidth(2) = 3500
    .ColWidth(3) = 900
    .ColWidth(4) = 0
    .ColWidth(5) = 1000
    .ColWidth(6) = 1200
    
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(1) = flexDTString
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColComboList(0) = cStrStore
End With
'DocTable.MoveLast
'xDoc_No.Text = DocTable.DOC_NO
'DocValid
CmdNewInv_Click
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub OLDItemInv_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If ItemInv.Row + 1 = ItemInv.Rows And ItemInv.Col = 0 And (ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = "") Then
        If Row > 1 Then ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = ItemInv.TextMatrix(ItemInv.Rows - 2, 0)
        ItemInv.Cell(flexcpForeColor, ItemInv.Rows - 1, 5, ItemInv.Rows - 1) = RGB(255, 0, 0)
    End If
End Sub
Private Sub OLDItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 3 Or ItemInv.Col = 4 Or ItemInv.Col = 5 Or ItemInv.Col = 6 Then KeyAscii = RetNumber(KeyAscii, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
If publicFlag = 1 Then
If UnInvTable.RecordCount > 0 Then
'    MsgBox "ÌÊÃœ »Ê‰«  €Ì— „”Ã·…"
End If
End If
End Sub

Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 46
    If MsgBox("Õ–› «·’‰› „‰ «·›« Ê—… ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        DocTable.Seek "=", xDoc_No.Text, ItemInv.TextMatrix(ItemInv.Row, 1)
        If Not DocTable.NoMatch Then DocTable.Delete
        ItemInv.RemoveItem ItemInv.Row
        Fillgrd
    End If
Case 112 And False
    If ItemInv.Col = 1 Then
        If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
            ChargesLookup
        Else
            ItemsLookup
        End If
    End If
End Select
End Sub
Private Sub OLDItemInv_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case KeyCode
Case 112
    If ItemInv.Col = 1 Then
        If ItemInv.TextMatrix(ItemInv.Row, 0) = "SS" Then
            ChargesLookup
        Else
            ItemsLookup
        End If
    End If
End Select
End Sub
Private Sub OLDItemInv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With ItemInv
Select Case ItemInv.Col
    Case 0
         If .EditText = "" Then
            MsgBox "ÌÃ»  ”Ã· «·„Œ“‰ "
            Cancel = True
        End If

    Case 1
        If ItemInv.TextMatrix(ItemInv.Row, 0) <> "SS" Then
            If .TextMatrix(.Row, 0) = "" Then
                MsgBox " ÕœÌœ «·„Œ“‰ √Ê·«"
                Cancel = True
                Exit Sub
            End If
            itemTable.FindFirst " ITEM = " & MyParn(ItemInv.EditText)
            If itemTable.NoMatch Then
                cMess = " ﬂÊœ «·’‰› €Ì— „”Ã· "
                Cancel = True
                Exit Sub
            Else
                DocTable.Seek "=", xDoc_No.Text, .EditText
                If Not DocTable.NoMatch Then
                    MsgBox "«·’‰› „”Ã· „‰ ﬁ»· ›Ï «·›« Ê—…"
                    Cancel = True
                    Exit Sub
                End If
                ItemInv.TextMatrix(ItemInv.Row, 2) = itemTable.desca
                Select Case publicFlag
                    Case 1
                        Select Case xLevel.Text
                            Case "0"
                                .TextMatrix(.Row, 5) = TurnValue(itemTable.COST1, Null, 0)
                            Case "1"
                                .TextMatrix(.Row, 5) = TurnValue(itemTable.COST2, Null, 0)
                            Case "3"
                                .TextMatrix(.Row, 5) = TurnValue(itemTable.price, Null, 0)
                            Case "2"
                                .TextMatrix(.Row, 5) = TurnValue(itemTable.COST4, Null, 0)
                        End Select
                    Case 2
                        SalesTable.Seek "=", xCode.Text, .EditText
                        If Not SalesTable.NoMatch Then
                            .TextMatrix(.Row, 5) = SalesTable.price
                        Else
                            MsgBox "·« ÌÊÃœ „»Ì⁄«  ·Â–« «·’‰› ··⁄„Ì·"
                            Cancel = True
                            Exit Sub
                        End If
                    Case 3
                        .TextMatrix(.Row, 5) = itemTable.COST1
                    Case 4
                        .TextMatrix(.Row, 5) = itemTable.COST1
                End Select
            End If
        Else
            ChargesTable.FindFirst "CODE = " & MyParn(ItemInv.EditText)
            If ChargesTable.NoMatch Then
                cMess = "·‰ Ì „  ”ÃÌ· «·„’—Ê›" & ItemInv.EditText & " ﬂÊœ «·„’—Ê› €Ì— „”Ã· "
                Cancel = True
            Else
                ItemInv.TextMatrix(ItemInv.Row, 2) = ChargesTable.desca
            End If
        End If
    Case 3
        If publicFlag = "1" Then
        End If
End Select
End With
End Sub

Private Sub xDISC_GotFocus()
    If TurnValue(ClientTable.DISC, Null, 0) <> 0 Then
        xDisc.Text = Format(Val(xTotItem.Text) * ClientTable.DISC / 100, "#0.00")
    End If
    xDisc.SelStart = 0
    xDisc.SelLength = Len(xDisc.Text)
End Sub
Private Sub xCODE_GotFocus()
    xCode.SelStart = 0
    xCode.SelLength = Len(xCode.Text)
End Sub

Private Sub xMAN_GotFocus()
    xMan.SelStart = 0
    xMan.SelLength = Len(xMan.Text)
End Sub
Private Sub xCash_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Val(xCash.Text) = (Val(xTotItem.Text) - Val(xDisc.Text)) Then
'        CmdSave_Click
'        If Not lManger Then CmdNewInv_Click
'    Else
'        SendKeys "{TAB}"
'    End If
'End If
End Sub
Private Sub xDisc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If publicFlag = "1" And ClientTable!CASH Then
            Load PaySal
            PaySal.xVisa.Text = ""
            PaySal.xCash.Text = Format(Val(xTotItem.Text) - Val(xDisc.Text), "#0.00")
                PaySal.Show 1
            If Str(Val(xCash.Text) + Val(xVisa.Text)) = Str(Val(xTotItem.Text) - Val(xDisc.Text)) Then
                cmdSave_Click
                CmdNewInv_Click
            Else
                MsgBox "—«Ã⁄ ﬁÌ„… «·”œœ"
            End If
        Else
            cmdSave_Click
            CmdNewInv_Click
        End If
    End If
End Sub
Private Sub xItem_Change()
'    If Len(xItem.Text) = 6 Then SendKeys "{TAB}"
End Sub
Private Sub xItem_GotFocus()
If xDoc_No.Text <> "" And xCode.Text <> "" Then
    xDoc_No.Enabled = False
    xCode.Enabled = False
    xMan.Enabled = False
    xDate.Enabled = False
    xDoc.Enabled = False
End If
End Sub
Private Sub xItem_KeyPress(KeyAscii As Integer)
    If KeyCode = 112 Then ItemsLookup
    If KeyAscii = 13 Then
        If xItem.Text = "" And Val(xTotItem.Text) <> 0 Then
            xPrice.Enabled = False
            xQty.Enabled = False
            
            xDisc.Enabled = True
        End If
    End If
End Sub
Private Sub xPrice_Validate(Cancel As Boolean)
'If Val(xPrice.Text) <= 0 Then
'    Cancel = True
'End If
End Sub
Private Sub xQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(xQty.Text) = 0 Then
        xQty.Text = ""
        Exit Sub
    End If
    If Val(xQty.Text) <> 0 And xDescItem.Caption <> "" And xItem.Text <> "" Then
        DocTable.AddNew
        DocTable.doc_no = xDoc_No.Text
        DocTable.doc = TurnValue(xDoc.Text, "", Null)
        DocTable.[Date] = xDate.Text
        DocTable.[Time] = Time
        DocTable.CODE = xCode.Text
        DocTable.Store = TurnValue(xStore.BoundText, "", Null)
        DocTable.Item = xItem.Text
        DocTable.price = Val(xPrice.Text)
        If publicFlag = 1 Then
            ItemCode.FindFirst " ITEM = " & MyParn(xItem.Text)
            If Not ItemCode.NoMatch Then DocTable.COST = ItemCode.COST
        End If
        DocTable.Quant = Val(xQty.Text)
        DocTable.total = Val(xPrice.Text) * Val(xQty.Text)
        DocTable.User = cUserName
        DocTable.Update
        Fillgrd
    End If
    xStore.Enabled = False
'    SendKeys "{TAB}"
    xPrice.Text = ""
    xQty.Text = ""
    xItem.Text = ""

End If
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    If publicFlag = 1 Or publicFlag = 2 Then
        Generalarray(2) = "Select Code As ﬂÊœ , DescA As ≈”„ From File3_10"
    Else
        Generalarray(2) = "Select Code As ﬂÊœ,DescA As ≈”„ From File4_10"
    End If
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xMAN_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xMan.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As ﬂÊœ , DescA As ≈”„ From File1_70 WHERE FLAG = 6 "
        
    GrdArray(1) = 1000
    GrdArray(2) = 3000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End If
End Sub
Private Sub xCode_Validate(Cancel As Boolean)
If xCode.Text = "" Then Cancel = True
xCode.Text = RetZero(xCode.Text, 6)
ClientTable.FindFirst "Code = " & MyParn(xCode.Text)
If Not ClientTable.NoMatch Then
    xclientDescA.Caption = ClientTable.desca
    If publicFlag = "1" Then
        xLevel.Text = TurnValue(ClientTable!price, Null, 0)
    Else
        xLevel.Text = "0"
    End If
Else
    Cancel = True
End If
End Sub
Private Sub xMAN_Validate(Cancel As Boolean)
If xMan.Text = "" Then Cancel = True
FlagTable.FindFirst "FLAG = 6 AND Code = " & MyParn(xMan.Text)
If Not FlagTable.NoMatch Then
    xManDescA.Caption = FlagTable.desca
Else
    Cancel = True
End If
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
    If Not IsDate(xDate.Text) Then Cancel = True
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And bopt1 Then
    xDoc_No.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    If publicFlag = 1 Or publicFlag = 2 Then
        Generalarray(2) = "Select " & myFileName & ".Doc_No as [«·„”·”·],format([Date],'dd-mm-yyyy') as [ «—ÌŒ «·›« Ê—…],File3_10.DescA as [«”„ «·„Ê—œ] " & _
                          " From " & myFileName & " Inner join file3_10 on " & myFileName & ".Code = file3_10.code "
        Generalarray(3) = " Where File3_10.DescA Like '*cFilter*' or doc_no Like '*cFilter*'"
        Generalarray(4) = " Group by " & myFileName & ".Doc_No,[date],File3_10.desca ORDER BY [DATE]"
    Else
        Generalarray(2) = "Select " & myFileName & ".Doc_No as [«·„”·”·],format([Date],'dd-mm-yyyy') as [ «—ÌŒ «·›« Ê—…],File4_10.DescA as [«”„ «·„Ê—œ] " & _
                          " From " & myFileName & " Inner join file4_10 on " & myFileName & ".Code = file4_10.code "
        Generalarray(3) = "Where File4_10.DescA Like '*cFilter*' or doc_no Like '*cFilter*' "
        Generalarray(4) = " Group by " & myFileName & ".Doc_No,[date],File4_10.desca ORDER BY [DATE]"
    End If
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
If publicFlag = 1 Then
    If KeyCode = 113 Then FrmUnInv.Show 1
End If
End Sub
Sub ChargesLookup()
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code as [«·„”·”·],DescA as [«·«”„] From file8_70 "
    Generalarray(3) = "WHERE DescA Like '*cFilter*'"
    GrdArray(1) = 1000
    GrdArray(2) = 4000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Function StrStore()
If storeTable.RecordCount > 0 Then
    storeTable.MoveFirst
    i = 1
    StrStore = "#" & storeTable!CODE & ";" & storeTable.desca
    storeTable.MoveNext
    Do While True
        i = i + 1
        If storeTable.EOF Then Exit Do
        StrStore = StrStore & "|#" & storeTable!CODE & ";" & storeTable.desca
        storeTable.MoveNext
    Loop
    StrStore = StrStore & "|#SS" & ";„’«—Ì›"
End If
End Function
Function myDelete()
    ' Õ–›  «·›« Ê—…
    cString = " DELETE  " & myFileName & " .* FROM " & myFileName & " WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
       
    ' Õ–› Õ—ﬂ… √’‰«› «·›« Ê—…
    cString = " DELETE  FILE1_11.* FROM FILE1_11 WHERE FILE1_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(itemMoveType)
    mydb.Execute cString
        
    ' Õ–› Õ—ﬂ… ⁄„Ì· «·›« Ê—…
    Select Case publicFlag
        Case 1
            cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
            cString = " DELETE  FILE6_22.* FROM FILE6_22 WHERE FILE6_22.DOC_no = " & MyParn(xDoc_No.Text)
        Case 2
            cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '5' "
        Case 3
            cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '4' "
        Case 4
            cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = '5' "
    End Select
    mydb.Execute cString
End Function
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = RetZero(xDoc_No.Text)
DocValid
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
'DocValid
End Sub
Private Sub CmdFirst_Click()
On Error Resume Next
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.MoveFirst
xDoc_No.Text = DocTable.doc_no
DocValid
End Sub
Private Sub CmdLast_Click()
'On Error Resume Next
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.MoveLast
xDoc_No.Text = DocTable.doc_no
DocValid
End Sub
Private Sub CmdNext_Click()
'On Error Resume Next
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.Seek ">=", xDoc_No.Text
If DocTable.NoMatch Then Exit Sub
Do Until DocTable.EOF
    If DocTable!doc_no <> xDoc_No.Text Then Exit Do
    DocTable.MoveNext
Loop
'DocTable.MoveNext
If Not DocTable.EOF Then
    xDoc_No.Text = DocTable.doc_no
    DocValid
End If
End Sub
Private Sub CmdPrevious_Click()
'On Error Resume Next
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.Seek "<=", xDoc_No.Text
If DocTable.NoMatch Then Exit Sub
'If DocTable.doc_no <> xDoc_No.Text Then Exit Sub
Do Until DocTable.BOF
    If DocTable.doc_no <> xDoc_No.Text Then Exit Do
    DocTable.MovePrevious
Loop
xDoc_No.Text = DocTable.doc_no
DocValid
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ItemsLookup
    If KeyCode = 113 Then
        If ItemInv.Rows = 2 And ItemInv.TextMatrix(1, 2) = "" Then
            MsgBox "·« ÌÊÃœ √’‰«› „”Ã·…"
        Else
            If MsgBox("”Ê› Ì „  ⁄·Ìﬁ «·»Ê‰", vbOKCancel) = vbOK Then
                If Not MYVALID Then Exit Sub
                If Not MyReplace Then Exit Sub
                
                UnInvTable.Seek "=", xDoc_No.Text
                If Not UnInvTable.NoMatch Then
                    UnInvTable.Edit
                Else
                    UnInvTable.AddNew
                End If
                UnInvTable.doc_no = xDoc_No.Text
                UnInvTable.Date = Date
                UnInvTable.Time = Time
                UnInvTable.Update
                
                CmdNewInv_Click
            End If
        End If
    End If
End Sub
Private Sub xItem_Validate(Cancel As Boolean)
    If Trim(xItem.Text) <> "" Then xItem.Text = RetZero(xItem.Text, 6)
    If xStore.Text = "" Then
        Cancel = False
        Exit Sub
    End If
    
    If xItem.Text = "" Then
        Cancel = False
        Exit Sub
    End If
    itemTable.FindFirst " ITEM = " & MyParn(xItem.Text)
    If itemTable.NoMatch Then
        Cancel = True
    Else
        Cancel = False
        xDescItem.Caption = itemTable.desca
        If publicFlag = 1 Then xBalItemInv.Caption = "«·—’Ìœ " & BalNoItem(xItem, MoveItemTable, xStore.BoundText)
        Select Case xLevel.Text
            Case "0"
                xPrice.Text = itemTable.COST1
            Case "1"
                xPrice.Text = itemTable.COST2
                Case "2"
                xPrice.Text = itemTable.COST4
            Case "3"
                xPrice.Text = itemTable.price
        End Select
    End If
End Sub
Private Sub xLevel_Validate(Cancel As Boolean)
    If xLevel.Text = "" Or xLevel.Text = "1" Or xLevel.Text = "2" Or xLevel.Text = "3" Or xLevel.Text = "" Or xLevel.Text = "4" Then
    Else
        Cancel = True
    End If
End Sub
Private Sub ViewClose()
    If DocTable.POSTED Then
        CmdDelInv.Enabled = False
        CmdSave.Enabled = False
    Else
        CmdDelInv.Enabled = True
        CmdSave.Enabled = True
    End If
End Sub
Private Sub xQty_GotFocus()
xQty.Text = 1
xQty.SelStart = 0
xQty.SelLength = 1
End Sub
Private Sub xPRICE_GotFocus()
xPrice.SelStart = 0
xPrice.SelLength = Len(xPrice.Text)
End Sub
Private Sub xQty_Validate(Cancel As Boolean)
'   On Error Resume Next
    
End Sub
Private Sub CmdAddPrint_Click()
On Error Resume Next
If publicFlag = "1" Then
    If bOpt5 Then
        TDaySal.Show 1
    Else
        TDaySal2.Show 1
    End If
    Exit Sub
End If
Dim tPrint As Recordset
    If Me.ItemInv.Rows = 1 Then Exit Sub
    Set tPrint = mydb.OpenRecordset("Select File7_20.Item,File1_10.DescA,File1_10.Price,File7_20.Quant,DOC_NO From File7_20 Inner join File1_10 on File7_20.Item = File1_10.item Where Doc_No = " & MyParn(xDoc_No.Text) & _
                                   " AND STORE <> 'zz' AND STORE <> 'SS' Order by File7_20.Item ")
    Set addPrint = mydb.OpenRecordset("addPrint", dbOpenDynaset)
    addPrint.FindFirst "DOC_NO = " & MyParn(tPrint.doc_no)
    If Not addPrint.NoMatch Then
        If MsgBox("«·›« Ê—…  „  —ÕÌ·Â« „‰ ﬁ»· .. «·€«¡ «·«’‰«› «·„—Õ·… ··ÿ»«⁄…", vbOKCancel + vbDefaultButton2, " —ÕÌ· ··ÿ»«⁄…") = vbCancel Then
            Exit Sub
        End If
    End If
    
    mydb.Execute "delete *  from addprint where Doc_no = " & MyParn(tPrint.doc_no)
    addPrint.Requery
    Do
        If TurnValue(tPrint!Item, Null, "") <> "" Then
            addPrint.FindFirst "Item = " & MyParn(tPrint.Item)
            If addPrint.NoMatch Then
                addPrint.AddNew
            Else
                addPrint.Edit
            End If
            addPrint!doc_no = tPrint!doc_no
            addPrint!Quant = TurnValue(addPrint!Quant, Null, 0) + tPrint!Quant
            addPrint!Item = tPrint!Item
            addPrint!isPrint = True
            addPrint.Update
            addPrint.Requery
        End If
        tPrint.MoveNext
    Loop Until tPrint.EOF
End Sub

Private Sub xVisa_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    CmdSave_Click
'    If Not lManger Then CmdNewInv_Click
'End If
End Sub
Private Sub myprint()
On Error Resume Next
With vp
     vp = " "
'   .Device = "GP-80160II"
    .Device = pDevice
    .StartDoc
   .TextAlign = taRightMiddle
   .TextColor = vbBlack
    
    .Width = 3500
    .Height = 3000
    .MarginLeft = 100
    .MarginRight = 100
    .MarginTop = 10
    .MarginFooter = 10
    .MarginBottom = 100
    
    .FontSize = 14
    .FontName = "Arial"
    .FontBold = True
    .TextAlign = taCenterMiddle
    .Paragraph = "** „œÌ‰… «·√Õ·«„ - ·Ê—«‰ **"
    
    .FontSize = 12
    .FontSize = 10
    .TextAlign = taRightMiddle
    .Paragraph = "»‹‹Ê‰ —ﬁ„ : " & xDoc_No.Text
    .Paragraph = " «—Ì‹‹‹‹Œ : " & Format(xDate.Text, "DD-MM-YYYY") & "      Êﬁ  : " & Format(Time, "SHORT TIME")
    .Paragraph = "»«∆‹‹‹‹‹⁄ : " & xManDescA.Caption & "   ===> " & Me.xStore.Text
    .Paragraph = " "
    
    .FontName = "Arial"
    .FontSize = 9
    .FontBold = False
    .StartTable
    .TableCell(tcRows) = ItemInv.Rows + 2
    .TableCell(tcCols) = 5
    .TableCell(tcColWidth, 1, 5) = 800
    .TableCell(tcColWidth, 1, 4) = 2000
    .TableCell(tcColWidth, 1, 3) = 500
    .TableCell(tcColWidth, 1, 2) = 0
    .TableCell(tcColWidth, 1, 1) = 1000
    .BorderStyle = bsNone
    
    .TableBorder = tbTopBottom
    .TableCell(tcRowBorderAbove, 1, 1, 1, 5) = 4

    For i = 0 To ItemInv.Rows - 1
        If ItemInv.TextMatrix(i, 0) <> "" Then
            .TableCell(tcText, i + 1, 5) = ItemInv.TextMatrix(i, 1)
            .TableCell(tcText, i + 1, 4) = Mid(ItemInv.TextMatrix(i, 2), 1, 14)
            .TableCell(tcText, i + 1, 3) = Format(ItemInv.TextMatrix(i, 3), "#")
            .TableCell(tcText, i + 1, 2) = Format(ItemInv.TextMatrix(i, 5), "FIXED")
            .TableCell(tcText, i + 1, 1) = Format(ItemInv.TextMatrix(i, 6), "FIXED")
        End If
    Next
    i = i + 1
    .TableCell(tcRowBorderBelow, 1, 1, 1, 5) = 4
    .TableCell(tcRowBorderAbove, i + 1, 1, i + 1, 5) = 4

'    .FontSize = 10
'    .FontBold = False

    .TableCell(tcText, i + 1, 4) = " ≈Ã„«·Ï"
    .TableCell(tcText, i + 1, 3) = xQ.Caption
    .TableCell(tcText, i + 1, 1) = Format(Val(xTotItem.Text), "FIXED")
    .TableCell(tcFontUnderline, i + 1, 1) = True
    .TableCell(tcAlign, 1, 1, i + 1, 5) = taRightMiddle
    
    If Val(xDisc.Text) <> 0 Then
        .TableCell(tcRows) = .TableCell(tcRows) + 2
        .TableCell(tcText, i + 2, 4) = " Œ’„"
        .TableCell(tcText, i + 2, 1) = Format(xDisc.Text, "FIXED")

        .TableCell(tcText, i + 3, 4) = " ≈Ã„«·Ï »⁄œ «·Œ’„"
        .TableCell(tcText, i + 3, 1) = Format(Val(xTotItem.Text) - Val(xDisc.Text), "FIXED")
        .TableCell(tcFontUnderline, i + 3, 1) = True
        .TableCell(tcAlign, 1, 1, i + 3, 5) = taRightMiddle
    End If
    .EndTable
    If Val(xVisa.Text) > 0 Then
    .Paragraph = "”œ«œ ›Ì“« : " & Format(xVisa.Text, "#0.00")
    .Paragraph = ""
    End If
    .FontSize = 8
    .FontBold = False
    .TextAlign = taCenterMiddle
    .Paragraph = ""
    .Paragraph = " *** ‘ﬂ—« ·“Ì«— ﬂ„ *** "
    .EndDoc
End With
Exit Sub
'myerror:
'MsgBox Err.Description
'MsgBox "Try Again This Order "
'Err.Clear
'Unload Me
End Sub
