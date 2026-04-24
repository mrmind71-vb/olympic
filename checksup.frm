VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form checksupfrm 
   Caption         =   "√Ê—«ﬁ œ›⁄"
   ClientHeight    =   6375
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   13200
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
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xName 
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
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1575
      Width           =   4290
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5535
      Width           =   5505
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "„—›Ê÷…"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "„Õ’·…"
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
         Index           =   2
         Left            =   2700
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Width           =   990
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "€Ì— „Õ’·…"
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
         Index           =   0
         Left            =   4095
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         Appearance      =   0  'Flat
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   3
         Left            =   180
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5535
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   48
         TabStop         =   0   'False
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
         Picture         =   "checksup.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "checksup.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   49
         TabStop         =   0   'False
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
         Picture         =   "checksup.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "checksup.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   50
         TabStop         =   0   'False
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
         Picture         =   "checksup.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "checksup.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   51
         TabStop         =   0   'False
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
         Picture         =   "checksup.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "checksup.frx":EB26
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "checksup.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Picture         =   "checksup.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1215
         MaskColor       =   &H00FFFFFF&
         Picture         =   "checksup.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Picture         =   "checksup.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Picture         =   "checksup.frx":1A6FA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Picture         =   "checksup.frx":1CC73
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.TextBox xdesca 
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
      Height          =   1050
      Left            =   90
      MaxLength       =   50
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2790
      Width           =   4290
   End
   Begin VB.CheckBox xOld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "‘Ìﬂ ”«»ﬁ"
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
      Left            =   6390
      RightToLeft     =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2115
   End
   Begin VB.TextBox xCode2 
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
      Left            =   9720
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1890
      Width           =   1230
   End
   Begin VB.TextBox xCode1 
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
      Left            =   9720
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1485
      Width           =   1230
   End
   Begin VB.TextBox xDATE_1 
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
      Left            =   8235
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2295
      Width           =   2715
   End
   Begin VB.TextBox xBANK_REC 
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
      Height          =   375
      Left            =   90
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2385
      Width           =   4290
   End
   Begin VB.TextBox xNAME4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   90
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1980
      Width           =   4290
   End
   Begin VB.TextBox xValue 
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
      Left            =   8235
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3105
      Width           =   2715
   End
   Begin VB.TextBox XSER_NO 
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
      Left            =   8595
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   765
      Width           =   2355
   End
   Begin VB.TextBox XCHK_ID 
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
      Left            =   8595
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   2355
   End
   Begin VB.TextBox xDATE_R 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8235
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2700
      Width           =   2715
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -2115
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
   Begin MSDataListLib.DataCombo XID_BANK 
      Height          =   315
      Left            =   8235
      TabIndex        =   7
      Top             =   3510
      Width           =   2715
      _ExtentX        =   4789
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1710
      Top             =   -45
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   -2115
      Top             =   990
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
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3960
      Width           =   12255
      Begin VB.TextBox xDATE_3 
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
         Height          =   360
         Left            =   7515
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   630
         Width           =   1890
      End
      Begin VB.TextBox xMEMO 
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
         Height          =   375
         Left            =   2745
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1035
         Width           =   6675
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "‘Ìﬂ €Ì— „Õ’·"
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
         Index           =   0
         Left            =   10125
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ŸÂÌ— &  Õ’Ì· "
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
         Index           =   2
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   225
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "—›÷ / —œ «·‘Ìﬂ"
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
         Index           =   1
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   225
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   2745
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   675
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“«‰… :"
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
         Left            =   5625
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·”œ«œ/  ŸÂÌ— / —›÷ :"
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   675
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«·«”„ :"
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
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label xName2 
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
      Height          =   375
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1890
      Width           =   4065
   End
   Begin VB.Label xName1 
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
      Height          =   375
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1485
      Width           =   4065
   End
   Begin VB.Label LabelCode2 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·„Ê—œ :"
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
      Left            =   11025
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "»‰ﬂ «·”Õ» :"
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
      Left            =   11070
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label label20 
      AutoSize        =   -1  'True
      Caption         =   "«·»Ì«‰ :"
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
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   2835
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "„ŸÂ— „‰ :"
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
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2070
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ :"
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
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2475
      Width           =   1650
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ :"
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
      Left            =   11070
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2430
      Width           =   1320
   End
   Begin VB.Label LabelCode 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·⁄„Ì· :"
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
      Left            =   11025
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2025
      Width           =   885
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "«·ﬁÌ„… :"
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
      Left            =   11070
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·‘Ìﬂ :"
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
      Left            =   11025
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "„”·”· ‘Ìﬂ :"
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
      Left            =   11025
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   810
      Width           =   1020
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ  Õ—Ì— :"
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
      Left            =   11070
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2835
      Width           =   990
   End
End
Attribute VB_Name = "checksupfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim oSearch As New Search3, oSearchClient As New Search3, oSearchSup As New Search3
Dim bSumMode As Boolean
Public sSer_no As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bEdit
cmdSave.Enabled = bEdit
CmdDel.Enabled = (nMode = LoadMode) And bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
XSER_NO.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 7)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = " SELECT FILE5_21.SER_NO,CASE WHEN (NOT CODE1 IS NULL) THEN '«·„Ê—œ : ' +  FILE4_10.DESCA ELSE CASE WHEN (NOT CODE2 IS NULL) THEN '«·⁄„Ì· : ' + FILE3_10.DESCA ELSE   [NAME] END END," & _
                  "CONVERT(VARCHAR(10),FILE5_21.DATE_1 ,111), CONVERT(VARCHAR(10),FILE5_21.DATE_r ,111), [VALUE] , CHK_ID" & _
                  "  From FILE5_21 LEFT JOIN FILE4_10 ON FILE5_21.CODE1 = FILE4_10.CODE LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE"
If retCloseOpt <> "3" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & " CLOSED = " & MyParn(retCloseOpt)
End If

Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ"
listarray(0, 1) = "(%%desca1%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"


listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "**[value]**"

GrdArray(0, 0) = "„”·”·"
GrdArray(0, 1) = 800

GrdArray(1, 0) = "‘Ìﬂ „‰"
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«” Õﬁ«ﬁ"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = " Õ—Ì—"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "ﬁÌ„…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "—ﬁ„ «·‘Ìﬂ"
GrdArray(5, 1) = 1200


searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„ «Ê—«ﬁ ﬁ»÷"
oSearch.Show 1
End Sub
Sub mydefine()
XSER_NO.Text = RetZero(Val(Newflag("FILE5_21", "ser_no")), 6)
XCHK_ID.Text = ""
xOld.Value = 0
xCode1.Text = ""
xCode2.Text = ""
XID_BANK.BoundText = ""
xName1.Caption = ""
xName2.Caption = ""
xNAME4.Text = ""
xBANK_REC.Text = ""
xDATE_1.Text = ""
xDATE_3.Text = ""
xDATE_R.Text = ""
XBOX.BoundText = ""
xValue.Text = ""
xMEMO.Text = ""
xDesca.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
Handlecontrols DefineMode
End Sub
Sub myDefine2()
XSER_NO.Text = RetZero(Val(Newflag("FILE5_21", "ser_no")), 5)
XCHK_ID.Text = Val(XCHK_ID.Text) + 1
xOld.Value = 0
xName2.Caption = ""
xNAME4.Text = ""
xDATE_1.Text = ""
xValue.Text = ""
xMEMO.Text = ""
xDesca.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xTransCode1.Text = ""
xTransCode2.Text = ""
xTransName1.Caption = ""
xTransName2.Caption = ""
xDateBank.Text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = xCode1.Name Then
    ActiveControl.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    xName1.Caption = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 1)
    Unload oSearchSup
    SendKeys "{TAB}"
ElseIf ActiveControl.Name = xCode2.Name Then
    ActiveControl.Text = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 0)
    xName2.Caption = oSearchClient.grid1.TextMatrix(oSearchClient.grid1.Row, 1)
    Unload oSearchClient
    SendKeys "{TAB}"
Else
    XSER_NO.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    myUndo
    Unload oSearch
End If
End Sub
Sub myload()
XSER_NO.Text = CardTable!Ser_no
XCHK_ID.Text = CardTable!CHK_ID & ""
xCode1.Text = CardTable!Code1 & ""
xCode2.Text = CardTable!Code2 & ""
xName1.Caption = CardTable!desca1 & ""
xName2.Caption = CardTable!Desca2 & ""
xNAME4.Text = CardTable!NAME4 & ""
xDesca.Text = CardTable!Desca & ""
xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xBANK_REC.Text = CardTable!Bank_rec & ""
xValue.Text = Format(CardTable!Value, "Fixed")
xDATE_1.Text = Format(CardTable!date_1, "YYYY-MM-DD")
xDATE_3.Text = Format(CardTable!date_3, "YYYY-MM-DD")
xDATE_R.Text = Format(CardTable!date_R, "YYYY-MM-DD")
xMEMO.Text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
XBOX.BoundText = CardTable!BOX & ""


Handlecontrols LoadMode
End Sub
Function MYVALID() As Boolean
If XSER_NO.Text = "" Then
    MsgBox "ÌÃ»  ”ÃÌ· „””·”· ··‘Ìﬂ"
    Exit Function
End If
If Not IsDate(xDATE_R.Text) Then
    MsgBox "ÌÃ»  ”ÃÌ·  «—ÌŒ «· Õ—Ì—"
    Exit Function
End If

If xCode1.Text <> "" Then
    If IsEmpty(GetField("select code from file4_10 where code = " & MyParn(xCode1.Text))) Then
        MsgBox "ﬂÊœ «·„Ê—œ €Ì— ’ÕÌÕ"
        Exit Function
    End If
End If

If xCode2.Text <> "" Then
    If IsEmpty(GetField("select code from file3_10 where code = " & MyParn(xCode2.Text))) Then
        MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
        Exit Function
    End If
End If


If XSER_NO.Enabled And Trim(xCode1.Text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_21 where code1 = " & MyParn(xCode1.Text) & " and Chk_Id = " & MyParn(XCHK_ID.Text)) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” " & "«·⁄„Ì· "
        Exit Function
    End If
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
On Error Resume Next
XCHK_ID.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
con.BeginTrans
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "delete  From FILE5_21 where Ser_No = " & MyParn(XSER_NO.Text)
End If
con.CommitTrans

CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "SER_NO < " & MyParn(XSER_NO.Text), , adSearchBackward, adBookmarkLast
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    If optclose(0).Value Then CmdAdd_Click Else mydefine
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
Private Sub cmdSave_Click()
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub cmdSum_Click()
myDefine2
On Error Resume Next
xDATE_1.SetFocus
Err.Clear
bSumMode = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
bEdit = True

openCon con

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_50 WHERE CODE > '500000'"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE5_10"

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub optclose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
openCardTable
myUndo
End Sub

Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XBOX.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox2.BoundText = ""
End Sub
Private Sub XCHK_ID_Validate(Cancel As Boolean)
If GetDesca("SELECT SER_NO FROM FILE5_21 WHERE SER_NO <> " & MyParn(XSER_NO.Text) & " AND  CHK_ID = " & MyParn(XCHK_ID.Text)) <> "" Then
    MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
End If
End Sub
Private Sub xClosed_Click(Index As Integer)
If Index = 0 Then
    xDATE_3.Text = ""
End If
End Sub
Private Sub xCode1_Change()
If xCode1.Text <> "" Then
    xCode2.Text = ""
    xCode2.Enabled = False
    xName2.Caption = ""
Else
    xCode2.Enabled = True
End If
End Sub
Private Sub xCode1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SupLookupAll Me, oSearchSup
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xName1.Caption = ""
If Trim(xCode1.Text) = "" Then Exit Sub
xCode1.Text = RetZero(xCode1.Text)
xName1.Caption = GetField("Select Desca from FILE4_10 where code = " & MyParn(xCode1.Text)) & ""
End Sub
Private Sub xCode2_Change()
If xCode2.Text <> "" Then
    xCode1.Text = ""
    xCode1.Enabled = False
    xName1.Caption = ""
Else
    xCode1.Enabled = True
End If
End Sub
Private Sub xCode2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookupAll Me, oSearchClient
End Sub

Private Sub xCODE2_LostFocus()
myLostFocus xCode2
xName2.Caption = ""
If Trim(xCode2.Text) = "" Then Exit Sub
xCode2.Text = RetZero(xCode2.Text)
xName2.Caption = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode2.Text))
End Sub

Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""
End Sub
Private Sub xSer_no_LostFocus()
myLostFocus XSER_NO
If Trim(XSER_NO.Text) = "" Then Exit Sub
XSER_NO.Text = RetZero(XSER_NO.Text, 6)
CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "CHK_ID", addstring(XCHK_ID.Text))
aInsert = AddFlag(aInsert, "OLD", xOld.Value)
aInsert = AddFlag(aInsert, "CODE1", addstring(xCode1.Text))
aInsert = AddFlag(aInsert, "CODE2", addstring(xCode2.Text))
aInsert = AddFlag(aInsert, "NAME", addstring(xName.Text))
aInsert = AddFlag(aInsert, "DESCA", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "BANK_REC", addstring(xBANK_REC.Text))
aInsert = AddFlag(aInsert, "DATE_1", addDate(xDATE_1.Text))
aInsert = AddFlag(aInsert, "DATE_3", addDate(xDATE_3.Text))
aInsert = AddFlag(aInsert, "DATE_R", addDate(xDATE_R.Text))
aInsert = AddFlag(aInsert, "[VALUE]", Val(xValue.Text))
aInsert = AddFlag(aInsert, "[NAME4]", addstring(xNAME4.Text))
aInsert = AddFlag(aInsert, "[BOX]", addstring(XBOX.BoundText))
aInsert = AddFlag(aInsert, "[MEMO]", addstring(xMEMO.Text))
aInsert = AddFlag(aInsert, "[ID_BANK]", addstring(XID_BANK.BoundText))
aInsert = AddFlag(aInsert, "[CLOSED]", addstring(retClose))
On Error GoTo myerror
con.BeginTrans
If XSER_NO.Enabled Then
    XSER_NO.Text = RetZero(Val(Newflag("FILE5_21", "ser_no")), 6)
    aInsert = AddFlag(aInsert, "SER_NO", addstring(XSER_NO.Text))
    con.Execute addInsert(aInsert, "FILE5_21")
Else
    con.Execute addUpdate(aInsert, "FILE5_21", "SER_NO = " & addstring(XSER_NO.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub handleopt()
For I = 1 To 3
    optclose(I).Enabled = GetDesca("select ser_no from " & cFileName & " where closed = " & MyParn(I)) <> ""
Next
End Sub
Private Function retClose() As String
Dim I As Integer
For I = 0 To 2
    If xClosed(I).Value Then
        retClose = I & ""
        Exit For
    End If
Next
End Function
Private Function retCloseOpt() As String
Dim I As Integer
For I = 0 To optclose.Count - 1
    If Me.optclose(I).Value Then
        retCloseOpt = I & ""
        Exit For
    End If
Next
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(XSER_NO.Text) <> "" Then
        CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "Select FILE5_21.*,FILE4_10.DESCA AS DESCA1,FILE3_10.DESCA AS DESCA2" & _
          " From (FILE5_21 LEFT JOIN FILE4_10 ON FILE5_21.CODE1 = FILE4_10.CODE) LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE"
If retCloseOpt <> "3" Then
    cString = cString & turn(cString) & " closed = " & MyParn(retCloseOpt)
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Private Sub xName_GotFocus()
myGotFocus xName
End Sub
Private Sub xName_LostFocus()
myLostFocus xName
End Sub
Private Sub xDateBank_GotFocus()
myGotFocus xDateBank
End Sub
Private Sub xDateBank_LostFocus()
myLostFocus xDateBank
myValidDate xDateBank
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub XCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xDATE_1_GotFocus()
myGotFocus xDATE_1
End Sub
Private Sub xDATE_1_LostFocus()
myLostFocus xDATE_1
myValidDate xDATE_1
End Sub
Private Sub xBANK_REC_GotFocus()
myGotFocus xBANK_REC
End Sub
Private Sub xBANK_REC_LostFocus()
myLostFocus xBANK_REC
End Sub
Private Sub xNAME4_GotFocus()
myGotFocus xNAME4
End Sub
Private Sub xNAME4_LostFocus()
myLostFocus xNAME4
End Sub
Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
Private Sub XSER_NO_GotFocus()
myGotFocus XSER_NO
End Sub
Private Sub XCHK_ID_GotFocus()
myGotFocus XCHK_ID
End Sub
Private Sub XCHK_ID_LostFocus()
myLostFocus XCHK_ID
End Sub
Private Sub xDATE_R_GotFocus()
myGotFocus xDATE_R
End Sub
Private Sub xDATE_R_LostFocus()
myLostFocus xDATE_R
myValidDate xDATE_R
End Sub
Private Sub XID_BANK_GotFocus()
myGotFocus XID_BANK
End Sub
Private Sub XID_BANK_LostFocus()
myLostFocus XID_BANK
If Not XID_BANK.MatchedWithList Then XID_BANK.BoundText = ""
End Sub
Private Sub xDATE_3_GotFocus()
myGotFocus xDATE_3
End Sub
Private Sub xDATE_3_LostFocus()
myLostFocus xDATE_3
myValidDate xDATE_3
End Sub
Private Sub xMEMO_GotFocus()
myGotFocus xMEMO
End Sub
Private Sub xMEMO_LostFocus()
myLostFocus xMEMO
End Sub
Private Sub xbox_GotFocus()
myGotFocus XBOX
End Sub
Private Sub xbox_LostFocus()
myLostFocus XBOX
If Not XBOX.MatchedWithList Then XBOX.BoundText = ""
End Sub
Private Sub xTransCode2_GotFocus()
myGotFocus xTransCode2
End Sub
Private Sub xTransCode1_GotFocus()
myGotFocus xTransCode1
End Sub
