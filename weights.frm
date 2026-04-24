VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form weightsfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·√’‰«›"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
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
   ScaleHeight     =   7560
   ScaleWidth      =   10050
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6885
      Width           =   3030
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   16
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
         Picture         =   "weights.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "weights.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   17
         Top             =   135
         Width           =   735
         _ExtentX        =   1296
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
         Picture         =   "weights.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "weights.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1530
         TabIndex        =   18
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
         Picture         =   "weights.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "weights.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2295
         TabIndex        =   19
         Top             =   135
         Width           =   690
         _ExtentX        =   1217
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
         Picture         =   "weights.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "weights.frx":EB26
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   2745
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   7215
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
         Picture         =   "weights.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2415
         MaskColor       =   &H00FFFFFF&
         Picture         =   "weights.frx":12FD8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "weights.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "weights.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4785
         MaskColor       =   &H00FFFFFF&
         Picture         =   "weights.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5985
         Picture         =   "weights.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc datastore 
      Height          =   330
      Left            =   7740
      Top             =   7425
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   630
      Width           =   9870
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
         Left            =   7515
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   1095
      End
      Begin VB.TextBox xItem 
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
         Height          =   360
         Left            =   6705
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton cmdSection 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   945
         Width           =   375
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   360
         Left            =   4815
         TabIndex        =   1
         Top             =   945
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   345
         Left            =   4815
         TabIndex        =   2
         Top             =   1350
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   345
         Left            =   4815
         TabIndex        =   23
         Top             =   1755
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
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
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì·"
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   630
         Width           =   480
      End
      Begin VB.Label xCode_Desca 
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   3075
      End
      Begin VB.Label Label4 
         Caption         =   "≈·Ì"
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "„‰"
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·”Ì«—…"
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "«·ﬂÊœ"
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
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.CommandButton CMD_FIX 
      BackColor       =   &H00DEE7D3&
      Caption         =   "Fix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9360
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   6885
      Top             =   7605
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5715
      Top             =   6840
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   7650
      Top             =   6840
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Left            =   855
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   -270
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   9870
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   3480
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   9690
         _cx             =   17092
         _cy             =   6138
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
   End
End
Attribute VB_Name = "weightsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim oSearch As New Search, oSearchItem As New Search
Const LoadMode = 1, DefineMode = 2
Dim nType As String
Dim con As New ADODB.Connection
Dim CardTable As New ADODB.Recordset
Private Sub cmdGroup_Click()
Dim nCode As String
nCode = xgroup.BoundText
itemsGroupFrm.Show
data2.Refresh
xgroup.BoundText = nCode
If Not xgroup.MatchedWithList Then xgroup.BoundText = ""
End Sub

Private Sub Command3_Click()

End Sub
End Sub
Private Sub cmdSection_Click()
Dim nCode As String
nCode = xSection.BoundText
sectionsfrm.bEdit = True
sectionsfrm.Show 1
data1.Refresh
xSection.BoundText = nCode
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
openCon con
CardTable.Open "SELECT * FROM FILE1_10 ORDER BY ITEM", con, adOpenKeyset, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE1_10SC ORDER BY DESCA "
Set xSection.RowSource = data1
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA "
Set xgroup.RowSource = data2
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon

'Set vsstore.DataSource = datastore
'datastore.ConnectionString = strCon

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdAdd_Click()
mydefine
xItem.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From FILE1_10  Where ITEM = " & MyParn(xItem.Text), nDelete
    DeleteSub
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "item < " & MyParn(xItem.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
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

Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·’‰› »‰Ã«Õ"
If xItem.Enabled Then
    CmdAdd_Click
Else
    CardTable.Requery
    CardTable.Find "ITEM = " & MyParn(xItem.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xItem.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "item = " & MyParn(xItem.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
ItemsLookupAll Me, oSearch, "FILE1_10.TYPE = 1"
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
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xItem.Enabled = Not (nMode = LoadMode)
End Sub
Sub mydefine()
xItem.Text = RetZero(Newflag("FILE1_10", "ITEM", con))
xDescA.Text = ""
xDescE.Text = ""
xgroup.BoundText = ""
xSection.BoundText = ""
XUNIT.BoundText = ""
xPrice.Text = ""
xCost.Text = ""
xShow.Value = 1
Handlecontrols DefineMode
Fixgrd
grid1.Rows = 1
grid1.AddItem ""
End Sub
Sub myload()
xItem.Text = CardTable!Item
xDescA.Text = CardTable!Desca & ""
xDescE.Text = CardTable!descE & ""
xSection.BoundText = CardTable!SECTION & ""
xgroup.BoundText = CardTable!Group & ""
xCost.Text = Myvalue(CardTable!Cost, "Fixed")
xPrice.Text = Myvalue(CardTable!price, "Fixed")
XUNIT.Text = CardTable!UNIT & ""
xShow.Value = IIf(CardTable!Show, 1, 0)
myloadgrd
'If xitem.Text <> "" Then
'    datastore.RecordSource = "select file0_40.desca , sum([in] - [out] ) as bal from file1_11 inner join file0_40 on file1_11.store = file0_40.code where file1_11.item = " & MyParn(xitem.Text) & " group by file1_11.store , file0_40.desca order by file1_11.store "
'    datastore.Refresh
'    FixGrdStore
'End If
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(8, 1)
aInsert(0, 0) = "Item"
aInsert(0, 1) = addstring(xItem.Text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDescA.Text)

aInsert(2, 0) = "[SECTION]"
aInsert(2, 1) = addvalue(xSection.BoundText)

aInsert(3, 0) = "[GROUP]"
aInsert(3, 1) = addvalue(xgroup.BoundText)

aInsert(4, 0) = "COST"
aInsert(4, 1) = Val(xCost.Text)

aInsert(5, 0) = "PRICE"
aInsert(5, 1) = Val(xPrice.Text)

aInsert(6, 0) = "[UNIT]"
aInsert(6, 1) = addstring(XUNIT.Text)

aInsert(7, 0) = "[DESCE]"
aInsert(7, 1) = addstring(xDescE.Text)

aInsert(8, 0) = "[SHOW]"
aInsert(8, 1) = xShow.Value

On Error GoTo myerror
con.BeginTrans

If xItem.Enabled Then
    con.Execute CreateInsert(aInsert, "FILE1_10")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_10", " WHERE FILE1_10.ITEM = " & MyParn(xItem.Text))
End If
myreplaceGrd
'myreplaceSub
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc(sControl)
On Error GoTo myerror
If sControl = grid1.Name Then
    nFound = FoundOtheritem(Row, Col, oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Exit Sub
    End If
    cItem = grid1.TextMatrix(grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    If Trim(cItem) <> Trim(grid1.TextMatrix(grid1.Row, 0)) Then calcRow grid1.Row, 0
    If grid1.Row = grid1.Rows - 1 Then
         grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.Select grid1.Rows - 1, 0
    End If
Else
    CardTable.Find "ITEM = " & oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    oSearch.Hide
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetKbLayout Lang_AR
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Unload Search
Set Search = Nothing
Err.Clear
End Sub
Private Sub Option1_Click(Index As Integer)
    'Cmd_Tree_Click
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 2 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRow(grid1.Row) Then
    grid1.AddItem "", grid1.Row
End If
If KeyCode = 112 And grid1.Col = 0 And grid1.Row <> 0 Then RawLookupAll Me, oSearchItem, grid1.Name
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And cmdSave.Enabled Then
    If MsgBox("Õ–› «·„ﬂÊ‰  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file1_30 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub xDescE_GotFocus()
SetKbLayout Lang_EN
End Sub
Private Sub xDescE_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub xitem_LostFocus()
If xItem.Text = "" Then Exit Sub
CardTable.Find "ITEM = " & MyParn(xItem.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
SetKbLayout Lang_AR
End Sub
Private Sub xCost4_GotFocus()
    If Val(xCost4.Text) = 0 Then xCost4.Text = Format(Val(xPrice.Text) / 1.2, "#0.00")
    xCost4.SelStart = 0
    xCost4.SelLength = Len(xCost4.Text)
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Cmd_Tree_Click
End Sub
Private Sub xgroup_LostFocus()
If Not xgroup.MatchedWithList Then
    xgroup.BoundText = ""
End If
End Sub
Function MYVALID() As Boolean
If xItem.Text = "" Then
    MsgBox "ﬂÊœ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "≈”„ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

        
If xgroup.BoundText = "" Then
    MsgBox "·«»œ „‰ «œ—«Ã „Ã„Ê⁄… ··’‰›"
    Exit Function
End If

MYVALID = True
End Function
Private Sub xUnit_Validate(Cancel As Boolean)
data4.Refresh
End Sub
Private Sub xNewItem_GotFocus()
SetKbLayout Lang_EN
xNewItem.SelStart = 0
xNewItem.SelLength = Len(xNewItem.Text)
End Sub
Private Sub xOldItem_GotFocus()
SetKbLayout Lang_EN
xOldItem.SelStart = 0
xOldItem.SelLength = Len(xOldItem.Text)
End Sub
Private Sub xCost_GotFocus()
xCost.SelStart = 0
xCost.SelLength = Len(xCost.Text)
End Sub
Private Sub xPrice_GotFocus()
xPrice.SelStart = 0
xPrice.SelLength = Len(xPrice.Text)
End Sub
Private Sub xPrice2_GotFocus()
xPrice2.SelStart = 0
xPrice2.SelLength = Len(xPrice2.Text)
End Sub
Private Sub xReorder_GotFocus()
xReorder.SelStart = 0
xReorder.SelLength = Len(xReorder.Text)
End Sub
Private Sub xShelf_GotFocus()
xShelf.SelStart = 0
xShelf.SelLength = Len(xShelf.Text)
End Sub
Private Sub xMaxDisc_GotFocus()
xMaxDisc.SelStart = 0
xMaxDisc.SelLength = Len(xMaxDisc.Text)
End Sub
Private Sub xdefsales_GotFocus()
xdefsales.SelStart = 0
xdefsales.SelLength = Len(xdefsales.Text)
End Sub
Private Sub xitem_GotFocus()
SetKbLayout Lang_EN
xItem.SelStart = 0
xItem.SelLength = Len(xItem.Text)
End Sub
Private Sub xDescA_GotFocus()
xDescA.SelStart = 0
xDescA.SelLength = Len(xDescA.Text)
End Sub
Private Sub xPackage_GotFocus()
xPackage.SelStart = 0
xPackage.SelLength = Len(xPackage.Text)
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 0 Then calcRow Row, Col
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then .AddItem ""
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then .RemoveItem OldRow
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then .RemoveItem .Row
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Trim(.TextMatrix(nRow, 0)) = "" Then Exit Function
If Not IsNumeric(.TextMatrix(nRow, 2)) Then Exit Function
If Val(.TextMatrix(nRow, 2)) <= 0 Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 0 Then
    If Trim(.EditText) = "" Then Cancel = True Else .EditText = RetZero(.EditText)
    If GetDesca("Select item from file2_10 where item = " & MyParn(.EditText)) = "" Then Cancel = True
End If

If Col = 2 Then
    If Not IsNumeric(.EditText) Then Cancel = True
End If
End With
End Sub
Private Sub myloadgrd()
With grid1
cString = "SELECT FILE1_20.ITEM_MAIN,FILE1_10.DESCA,QUANT,ID" & _
          " From FILE1_20 INNER JOIN FILE1_10 ON FILE1_20.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "FILE1_20.ITEM = " & MyParn(xItem.Text)
cString = cString & " Order by ROW"
data10.RecordSource = cString
data10.Refresh
.AddItem ""
Fixgrd
End With
End Sub
Private Sub Fixgrd()
With grid1
    .Cols = 4
    .FormatString = "«·ﬂÊœ|" & "«·’‰›|" & "«·ﬂ„Ì…|"
    .ColWidth(0) = 1200
    .ColWidth(1) = 4000
    .ColWidth(2) = 1200
    .ColHidden(.Cols - 1) = True
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub calcRow(Row, Col)
Dim nBalance As Double
grid1.TextMatrix(Row, 1) = ""
aRet = aGetDesca("select desca from FILE2_10 where item = " & MyParn(grid1.TextMatrix(Row, 0)))
If UBound(aRet) > 0 Then
    grid1.TextMatrix(Row, 1) = aRet(1) & ""
End If
'CalcTotals
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
Private Function myreplaceSub() As Boolean
Dim aInsert(6, 1)
aInsert(0, 0) = "Item"
aInsert(0, 1) = addstring(xItem.Text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDescA.Text & turn(xgroup.Text, " (") & xgroup.Text & turn(xgroup.Text & "", ")"))

aInsert(2, 0) = "[GROUP]"
aInsert(2, 1) = addvalue(xgroup.BoundText)

aInsert(3, 0) = "COST"
aInsert(3, 1) = Val(xCost.Text)

aInsert(4, 0) = "[UNIT]"
aInsert(4, 1) = addstring(XUNIT.Text)

aInsert(5, 0) = "[TYPE]"
aInsert(5, 1) = "2"

aInsert(6, 0) = "PRICE"
aInsert(6, 1) = Val(xPrice.Text)

If xItem.Enabled Then
    con.Execute CreateInsert(aInsert, "FILE2_10")
    con.Execute "insert into FILE1_30(ITEM,ITEMSUB,QUANT)" & _
                "VALUES(" & _
                  addstring(xItem.Text) & "," & _
                  addstring(xItem.Text) & "," & _
                  "1" & _
                ")"
Else
    con.Execute CreateUpdate(aInsert, "FILE2_10", " WHERE FILE2_10.ITEM = " & MyParn(xItem.Text))
End If
myreplaceSub = True
End Function
Private Sub DeleteSub()
    con.Execute "Delete  From FILE2_10  Where ITEM = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), nDelete
    con.Execute "Delete  From FILE1_30  Where ITEM = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), nDelete
End Sub
Private Sub myreplaceGrd()
Dim aInsert(3, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "ITEM"
        aInsert(0, 1) = addstring(xItem.Text)
               
        aInsert(1, 0) = "ITEMSUB"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = .TextMatrix(i, 2)

        aInsert(3, 0) = "row"
        aInsert(3, 1) = i
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE1_30")
        Else
            con.Execute CreateUpdate(aInsert, "FILE1_30", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub

