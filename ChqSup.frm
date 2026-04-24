VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chqsupfrm 
   Caption         =   "√Ê—«ﬁ ﬁ»÷"
   ClientHeight    =   5760
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9705
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
   ScaleHeight     =   5760
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   10
      Top             =   2745
      Width           =   3300
   End
   Begin VB.CheckBox xOld 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "‘Ìﬂ ”«»ﬁ"
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
      Height          =   240
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   675
      Width           =   2115
   End
   Begin VB.TextBox xName2 
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
      Left            =   2745
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1665
      Width           =   4020
   End
   Begin VB.TextBox xCode2 
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
      Left            =   6795
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   1230
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9705
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   9705
      Begin VB.CommandButton cmdSum 
         Caption         =   "‘Ìﬂ«  „Ã„⁄…"
         Height          =   375
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":0000
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton CmdExit 
         Height          =   375
         Left            =   3420
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":0532
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         Height          =   375
         Left            =   5265
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":067C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Õ›Ÿ"
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdInform 
         Height          =   375
         Left            =   7995
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":0BAE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   375
         Left            =   7080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":0FF0
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   375
         Left            =   6165
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":1522
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         Height          =   375
         Left            =   4335
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSup.frx":1FF3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9705
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5310
      Width           =   9705
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "«·ﬂ‹‹‹·"
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
         Index           =   3
         Left            =   3900
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   75
         Width           =   1215
      End
      Begin VB.CommandButton cmdfirst 
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
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   1305
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   855
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
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
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
         Top             =   45
         Width           =   435
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "€Ì— „Õ’·…"
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
         Index           =   0
         Left            =   7785
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   90
         Width           =   1215
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "„Õ’·…"
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
         Index           =   2
         Left            =   6570
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   90
         Width           =   990
      End
      Begin VB.OptionButton optclose 
         BackColor       =   &H80000010&
         Caption         =   "„—›Ê÷…"
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
         Index           =   1
         Left            =   5175
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.TextBox xCode1 
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
      Left            =   6795
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1305
      Width           =   1230
   End
   Begin VB.TextBox xNAME1 
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
      Left            =   2745
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1305
      Width           =   4020
   End
   Begin VB.TextBox xDATE_1 
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
      Left            =   5310
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2025
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
      Height          =   330
      Left            =   90
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2385
      Width           =   3300
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
      Height          =   330
      Left            =   90
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2025
      Width           =   3300
   End
   Begin VB.TextBox xValue 
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
      Left            =   5310
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2745
      Width           =   2715
   End
   Begin VB.TextBox XSER_NO 
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
      Left            =   5670
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   2355
   End
   Begin VB.TextBox XCHK_ID 
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
      Left            =   5670
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   2355
   End
   Begin VB.TextBox xDATE_R 
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
      Height          =   330
      Left            =   5310
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2385
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
      Left            =   5310
      TabIndex        =   7
      Top             =   3105
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   8700
      Top             =   3975
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
      Height          =   1410
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3825
      Width           =   7845
      Begin VB.TextBox xDATE_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
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
         Height          =   315
         Left            =   225
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   945
         Width           =   4965
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ìﬂ €Ì— „Õ’·"
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
         Height          =   345
         Index           =   0
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         Value           =   -1  'True
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   " ŸÂÌ— &  Õ’Ì· "
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
         Height          =   345
         Index           =   2
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1740
      End
      Begin VB.OptionButton xClosed 
         Alignment       =   1  'Right Justify
         Caption         =   "—›÷ / —œ «·‘Ìﬂ"
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
         Height          =   345
         Index           =   1
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   150
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   225
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“«‰… :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·”œ«œ/  ŸÂÌ— / —›÷"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   675
         Width           =   1860
      End
   End
   Begin VB.Label LabelCode2 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·⁄„Ì· :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﬂÊœ «·„Ê—œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   75
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "»‰ﬂ «·”Õ» :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3195
      Width           =   1005
   End
   Begin VB.Label label20 
      AutoSize        =   -1  'True
      Caption         =   "«·»Ì«‰ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2745
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "„ŸÂ— „‰ :"
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
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2070
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2430
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2115
      Width           =   1380
   End
   Begin VB.Label LabelCode 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·„Ê—œ :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "«·ﬁÌ„… :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2835
      Width           =   570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·‘Ìﬂ :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "„”·”· ‘Ìﬂ :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ  Õ—Ì— :"
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2430
      Width           =   945
   End
End
Attribute VB_Name = "chqsupfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bSumMode As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bEdit
cmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value) And bEdit
cmdSum.Enabled = (nMode = LoadMode Or optclose(0).Value) And bEdit
CmdUndo.Enabled = (nMode = LoadMode Or optclose(0).Value)
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

If optclose(0).Value Then cwhere = "CLOSED = '0'"
If optclose(1).Value Then cwhere = "CLOSED = '1'"
If optclose(2).Value Then cwhere = "CLOSED = '2'"

Set Generalarray(0) = Me
Generalarray(1) = " SELECT FILE5_21.SER_NO,FILE5_21.DESCA1,CONVERT(VARCHAR(10),FILE5_21.DATE_1 ,111), CONVERT(VARCHAR(10),FILE5_21.DATE_r ,111), [VALUE] , CHK_ID From FILE5_21 " & turn(cwhere, " where ") & cwhere
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
Search3.Caption = "«” ⁄·«„ «Ê—«ﬁ ﬁ»÷"
Search3.Show 1
End Sub
Sub CLIENTLOOKUP(Optional nFlag As Integer = 1)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
If nFlag = 1 Then
    Generalarray(1) = "SELECT Code , Desca  From file4_10"
Else
    Generalarray(1) = "SELECT Code , Desca  From file3_10"
End If
Generalarray(2) = " Order by Code"
Generalarray(3) = 4000
Generalarray(5) = False

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·«”„"
GrdArray(1, 1) = 3000

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%desca%%)"

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„" & " " & IIf(nFlag = 1, "«·„Ê—œÌ‰", "«·⁄„·«¡")
Search3.Show 1
End Sub
Sub mydefine()
XSER_NO.Text = RetZero(Val(Newflag("FILE5_21", "ser_no")), 5)
XCHK_ID.Text = ""
xOld.Value = 0
xCode1.Text = ""
xCode2.Text = ""
XID_BANK.BoundText = ""
xName1.Text = ""
xName2.Text = ""
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
xName2.Text = ""
xNAME4.Text = ""
xDATE_1.Text = ""
xValue.Text = ""
xMEMO.Text = ""
xDesca.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
Handlecontrols DefineMode
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
Else
    CardTable.Find "SER_NO = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Unload Search3
End If
End Sub
Sub myload()
XSER_NO.Text = CardTable!Ser_no
XCHK_ID.Text = CardTable!CHK_ID & ""
xCode1.Text = CardTable!Code1 & ""
xCode2.Text = CardTable!Code2 & ""
xName1.Text = CardTable!desca1 & ""
xName2.Text = CardTable!Desca2 & ""
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
    If GetDesca("select code from file4_10 where code = " & MyParn(xCode1.Text)) = "" Then Exit Function
End If

If xCode2.Text <> "" Then
    If GetDesca("select code from file3_10 where code = " & MyParn(xCode2.Text)) = "" Then Exit Function
End If

If XSER_NO.Enabled And Trim(xCode1.Text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_21 where code1 = " & MyParn(xCode1.Text) & " and Chk_Id = " & MyParn(XCHK_ID.Text)) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” " & " «·„Ê—œ "
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
'msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ »Ì«‰«  «·⁄„Ì· »‰Ã«Õ"
CardTable.Requery
If XSER_NO.Enabled Then
    If bSumMode Then cmdSum_Click Else CmdAdd_Click
Else
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then
            CardTable.Find "SER_NO < " & MyParn(XSER_NO.Text), , adSearchBackward, adBookmarkLast
            If CardTable.BOF Then CardTable.MoveFirst
        End If
        myload
    Else
        mydefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
bSumMode = False
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If XSER_NO.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "ser_no = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
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
Me.Caption = "√Ê—«ﬁ ﬁ»÷"
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_50"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE5_10"

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

optclose(0).Value = True
'If Not (CardTable.EOF And CardTable.BOF) Then
'    CardTable.MoveLast
'    MyLoad
'Else
'    CmdAdd_Click
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
closeCon con
End Sub

Private Sub optclose_Click(Index As Integer)
cString = "Select FILE5_21.*,FILE4_10.DESCA AS DESCA1,FILE3_10.DESCA AS DESCA2" & _
          " From (FILE5_21 LEFT JOIN FILE4_10 ON FILE5_21.CODE1 = FILE4_10.CODE) LEFT JOIN FILE3_10 ON FILE5_21.CODE2 = FILE3_10.CODE"
If Index <> 3 Then
    cString = cString & turn(cString) & " closed = " & MyParn(Index)
End If
cString = cString & " Order by Ser_No"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenKeyset, adLockOptimistic, adCmdText

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    If optclose(0).Value Then
        CmdAdd_Click
    Else
        mydefine
    End If
End If
End Sub
Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XBOX.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xBox2.BoundText = ""
End Sub

Private Sub XCHK_ID_Validate(Cancel As Boolean)
If publicFlag = 2 And XSER_NO.Enabled Then
    If GetDesca("SELECT SER_NO FROM FILE5_21 WHERE SER_NO <> " & MyParn(XSER_NO.Text) & " AND  CHK_ID = " & MyParn(XCHK_ID.Text)) <> "" Then
        MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
    End If
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
    xName2.Text = ""
Else
    xCode2.Enabled = True
End If
End Sub

Private Sub xCode1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xCode1_LostFocus()
xName1.Text = ""
If Trim(xCode1.Text) = "" Then Exit Sub
xCode1.Text = RetZero(xCode1.Text)
xName1.Text = GetDesca("Select Desca from FILE4_10 where code = " & MyParn(xCode1.Text))
End Sub
Private Sub xCode2_Change()
If xCode2.Text <> "" Then
    xCode1.Text = ""
    xCode1.Enabled = False
    xName1.Text = ""
Else
    xCode1.Enabled = True
End If
End Sub

Private Sub xCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub
Private Sub xCODE2_LostFocus()
If Trim(xCode2.Text) = "" Then Exit Sub
xName2.Text = ""
xCode2.Text = RetZero(xCode2.Text)
xName2.Text = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode2.Text))
End Sub
Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""
End Sub
Private Sub xSer_no_LostFocus()
If Trim(XSER_NO.Text) = "" Then Exit Sub
XSER_NO.Text = RetZero(XSER_NO.Text, 5)
CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function myreplace() As Boolean
Dim aInsert(16, 1)
aInsert(0, 0) = "SER_NO"
aInsert(0, 1) = addstring(XSER_NO.Text)

aInsert(1, 0) = "CHK_ID"
aInsert(1, 1) = addstring(XCHK_ID.Text)

aInsert(2, 0) = "OLD"
aInsert(2, 1) = xOld.Value

aInsert(3, 0) = "code1"
aInsert(3, 1) = addstring(xCode1.Text)

aInsert(4, 0) = "code2"
aInsert(4, 1) = addstring(xCode2.Text)

aInsert(5, 0) = "[desca]"
aInsert(5, 1) = addstring(xDesca.Text)

aInsert(6, 0) = "Bank_rec"
aInsert(6, 1) = addstring(xBANK_REC.Text)

aInsert(7, 0) = "date_1"
aInsert(7, 1) = addDate(xDATE_1.Text)

aInsert(8, 0) = "date_3"
aInsert(8, 1) = addDate(xDATE_3.Text)

aInsert(9, 0) = "date_r"
aInsert(9, 1) = addDate(xDATE_R.Text)

aInsert(10, 0) = "[VALUE]"
aInsert(10, 1) = Val(xValue.Text)

aInsert(11, 0) = "NAME4"
aInsert(11, 1) = addstring(xNAME4.Text)

aInsert(12, 0) = "BOX"
aInsert(12, 1) = addstring(XBOX.BoundText)

aInsert(13, 0) = "MEMO"
aInsert(13, 1) = addstring(xMEMO.Text)

aInsert(14, 0) = "ID_BANK"
aInsert(14, 1) = addstring(XID_BANK.BoundText)

If Trim(xCode1.Text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("„Ê—œ:" & xName1.Text)
ElseIf Trim(xCode2.Text) <> "" Then
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = addstring("⁄„Ì·:" & xName1.Text)
Else
    aInsert(15, 0) = "DESCA1"
    aInsert(15, 1) = "NULL"
End If

aInsert(16, 0) = "Closed"
aInsert(16, 1) = addstring(retClose)

On Error GoTo myerror
con.BeginTrans
If XSER_NO.Enabled Then
    XSER_NO.Text = RetZero(Val(Newflag("FILE5_21", "ser_no")), 5)
    aInsert(0, 1) = addstring(XSER_NO.Text)
    con.Execute CreateInsert(aInsert, "FILE5_21")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_21", " where SER_NO = " & addstring(XSER_NO.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function retClose() As String
Dim I As Integer
For I = 0 To 2
    If xClosed(I).Value Then
        retClose = I & ""
        Exit For
    End If
Next
End Function
Private Sub xDescA_GotFocus()
xDesca.SelStart = 0
xDesca.SelLength = Len(xDesca.Text)
End Sub
Private Sub xName2_GotFocus()
xName2.SelStart = 0
xName2.SelLength = Len(xName2.Text)
End Sub
Private Sub XCode2_GotFocus()
xCode2.SelStart = 0
xCode2.SelLength = Len(xCode2.Text)
End Sub
Private Sub xCode1_GotFocus()
xCode1.SelStart = 0
xCode1.SelLength = Len(xCode1.Text)
End Sub
Private Sub xNAME1_GotFocus()
xName1.SelStart = 0
xName1.SelLength = Len(xName1.Text)
End Sub
Private Sub xDATE_1_GotFocus()
xDATE_1.SelStart = 0
xDATE_1.SelLength = Len(xDATE_1.Text)
End Sub
Private Sub xBANK_REC_GotFocus()
xBANK_REC.SelStart = 0
xBANK_REC.SelLength = Len(xBANK_REC.Text)
End Sub
Private Sub xNAME4_GotFocus()
xNAME4.SelStart = 0
xNAME4.SelLength = Len(xNAME4.Text)
End Sub
Private Sub xValue_GotFocus()
xValue.SelStart = 0
xValue.SelLength = Len(xValue.Text)
End Sub
Private Sub XSER_NO_GotFocus()
XSER_NO.SelStart = 0
XSER_NO.SelLength = Len(XSER_NO.Text)
End Sub
Private Sub XCHK_ID_GotFocus()
XCHK_ID.SelStart = 0
XCHK_ID.SelLength = Len(XCHK_ID.Text)
End Sub
Private Sub xDATE_R_GotFocus()
xDATE_R.SelStart = 0
xDATE_R.SelLength = Len(xDATE_R.Text)
End Sub
Private Sub xDATE_3_GotFocus()
xDATE_3.SelStart = 0
xDATE_3.SelLength = Len(xDATE_3.Text)
End Sub
Private Sub xMEMO_GotFocus()
xMEMO.SelStart = 0
xMEMO.SelLength = Len(xMEMO.Text)
End Sub
