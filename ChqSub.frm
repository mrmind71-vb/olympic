VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chqClientfrm 
   Caption         =   "√Ê—«ﬁ ﬁ»÷"
   ClientHeight    =   6780
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9225
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
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xDateBank 
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
      Left            =   5130
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   3465
      Visible         =   0   'False
      Width           =   2715
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
      TabIndex        =   10
      Top             =   2745
      Width           =   3300
   End
   Begin VB.CheckBox xOld 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   3315
      RightToLeft     =   -1  'True
      TabIndex        =   19
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
      Left            =   2565
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   21
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
      Left            =   6615
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
      ScaleWidth      =   9225
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Width           =   9225
      Begin VB.CommandButton cmdSum 
         Caption         =   "‘Ìﬂ«  „Ã„⁄…"
         Height          =   375
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":0000
         RightToLeft     =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   2895
      End
      Begin VB.CommandButton CmdExit 
         Height          =   375
         Left            =   3420
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":0532
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         Height          =   375
         Left            =   5250
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":067C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Õ›Ÿ"
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdInform 
         Height          =   375
         Left            =   7995
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":0BAE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   375
         Left            =   7080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":0FF0
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   375
         Left            =   6165
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":1522
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         Height          =   375
         Left            =   4335
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ChqSub.frx":1FF3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   54
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
      ScaleWidth      =   9225
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6330
      Width           =   9225
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
         TabIndex        =   63
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
      Left            =   6615
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
      Left            =   2565
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   20
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
      Left            =   5130
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
      Left            =   5130
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
      Left            =   5490
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
      Left            =   5490
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
      Left            =   5130
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
      Left            =   5130
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
      TabIndex        =   36
      Top             =   3825
      Width           =   7845
      Begin VB.TextBox xDATE_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         Width           =   1740
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   225
         TabIndex        =   17
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
         TabIndex        =   46
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   675
         Width           =   1860
      End
   End
   Begin VB.Frame frmTrans 
      Height          =   1095
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   5175
      Width           =   7845
      Begin VB.TextBox xTransName2 
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
         Height          =   315
         Left            =   675
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   585
         Width           =   2340
      End
      Begin VB.TextBox xTransCode2 
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   585
         Width           =   1200
      End
      Begin VB.TextBox xTransCode1 
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
         Left            =   4365
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1200
      End
      Begin VB.TextBox xTransName1 
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
         Height          =   315
         Left            =   675
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   2340
      End
      Begin VB.Label xLab3 
         AutoSize        =   -1  'True
         Caption         =   " ŸÂÌ— «·‘Ìﬂ ≈·Ï «·⁄„Ì·"
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
         Left            =   5745
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label xLab4 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·⁄„Ì·"
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
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   690
         Width           =   690
      End
      Begin VB.Label xLab2 
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·„Ê—œ"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   225
         Width           =   705
      End
      Begin VB.Label xLab1 
         AutoSize        =   -1  'True
         Caption         =   " ŸÂÌ— «·‘Ìﬂ ≈·Ï «·„Ê—œ"
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
         Left            =   5775
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   1560
      End
   End
   Begin VB.Label lblDateBank 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·«Ìœ«⁄ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   3420
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label LabelCode2 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·„Ê—œ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1665
      Width           =   810
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
      TabIndex        =   34
      Top             =   75
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "»‰ﬂ «·«Ìœ«⁄ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3105
      Width           =   810
   End
   Begin VB.Label label20 
      AutoSize        =   -1  'True
      Caption         =   "«·»Ì«‰ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   2745
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "„ŸÂ— „‰ :"
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
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "«·»‰ﬂ «·„”ÕÊ» ⁄·ÌÂ :"
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
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2430
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2025
      Width           =   1185
   End
   Begin VB.Label LabelCode 
      AutoSize        =   -1  'True
      Caption         =   "ﬂÊœ «·⁄„Ì· :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1350
      Width           =   795
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "«·ﬁÌ„… :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2745
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·‘Ìﬂ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   990
      Width           =   780
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "„”·”· ‘Ìﬂ :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   630
      Width           =   915
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ  Õ—Ì— :"
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2340
      Width           =   945
   End
End
Attribute VB_Name = "chqClientfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim bSumMode As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bEdit
cmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value) And bEdit
cmdSum.Enabled = (nMode = LoadMode Or optclose(0).Value) And bEdit
CmdUndo.Enabled = (nMode = LoadMode Or optclose(0).Value)
CmdDel.Enabled = (nMode = LoadMode) And bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdfirst.Enabled = (nMode = LoadMode)
xTransCode1.Enabled = (optclose(2).Value)
xTransCode2.Enabled = (optclose(2).Value)
XSER_NO.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 7)
Dim GrdArray(6, 1)

If optclose(0).Value Then cWhere = "CLOSED = '0'"
If optclose(1).Value Then cWhere = "CLOSED = '1'"
If optclose(2).Value Then cWhere = "CLOSED = '2'"

Set Generalarray(0) = Me
Generalarray(1) = " SELECT FILE5_20.SER_NO,FILE5_20.DESCA1,CONVERT(VARCHAR(10),FILE5_20.DATE_1 ,111), CONVERT(VARCHAR(10),FILE5_20.DATE_r ,111), [VALUE] , CHK_ID, FILE5_20.TRANSNAME1 From FILE5_20 " & turn(cWhere, " where ") & cWhere
Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ"
listarray(0, 1) = "(%%desca1%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"


listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "##[value]##"

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

GrdArray(6, 0) = " ŸÂÌ—"
GrdArray(6, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ «Ê—«ﬁ ﬁ»÷"
Search3.Show 1
End Sub
Sub CLIENTLOOKUP(Optional nFlag As Integer = 1)
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
If nFlag = 1 Then
    Generalarray(1) = "SELECT Code , Desca  From file3_10"
Else
    Generalarray(1) = "SELECT Code , Desca  From file4_10"
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
Search3.Caption = "«” ⁄·«„" & " " & IIf(nFlag = 1, "«·⁄„·«¡", "«·„Ê—œÌ‰")
Search3.Show 1
End Sub
Sub myDefine()
XSER_NO.Text = RetZero(Val(Newflag("FILE5_20", "ser_no")), 5)
XCHK_ID.Text = ""
xOld.Value = 0
xCode1.Text = ""
xCode2.Text = ""
XID_BANK.BoundText = ""
xNAME1.Text = ""
xName2.Text = ""
xNAME4.Text = ""
xBANK_REC.Text = ""
xDATE_1.Text = ""
xDATE_3.Text = ""
xDATE_R.Text = ""
XBOX.BoundText = ""
xValue.Text = ""
xMEMO.Text = ""
xdesca.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xTransCode1.Text = ""
xTransCode2.Text = ""
xTransName1.Text = ""
xTransName2.Text = ""
If publicFlag = 1 Then xDateBank.Text = ""
Handlecontrols DefineMode
End Sub
Sub myDefine2()
XSER_NO.Text = RetZero(Val(Newflag("FILE5_20", "ser_no")), 5)
XCHK_ID.Text = Val(XCHK_ID.Text) + 1
xOld.Value = 0
xName2.Text = ""
xNAME4.Text = ""
xDATE_1.Text = ""
xValue.Text = ""
xMEMO.Text = ""
xdesca.Text = ""
xClosed(0) = True
xClosed(1) = False
xClosed(2) = False
xTransCode1.Text = ""
xTransCode2.Text = ""
xTransName1.Text = ""
xTransName2.Text = ""
If publicFlag = 1 Then xDateBank.Text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
Else
    CardTable.Find "SER_NO = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    MyLoad
    Unload Search3
End If
End Sub
Sub MyLoad()
XSER_NO.Text = CardTable!Ser_no
XCHK_ID.Text = CardTable!CHK_ID & ""
xCode1.Text = CardTable!Code1 & ""
xCode2.Text = CardTable!Code2 & ""
xNAME1.Text = CardTable!desca1 & ""
xName2.Text = CardTable!Desca2 & ""
xNAME4.Text = CardTable!NAME4 & ""
xdesca.Text = CardTable!Desca & ""
xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xBANK_REC.Text = CardTable!Bank_rec & ""
xValue.Text = Format(CardTable!Value, "Fixed")
xDATE_1.Text = Format(CardTable!date_1, "dd-mm-yyyy")
xDATE_3.Text = Format(CardTable!date_3, "dd-mm-yyyy")
xDATE_R.Text = Format(CardTable!date_R, "dd-mm-yyyy")
xDateBank.Text = Format(CardTable!DateBank, "dd-mm-yyyy")
xMEMO.Text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
XBOX.BoundText = CardTable!Box & ""

xTransCode1.Text = CardTable!TransCode1 & ""
xTransCode2.Text = CardTable!TRANSCODE2 & ""

xTransName1.Text = CardTable!TRANSNAME1 & ""
xTransName2.Text = CardTable!transname2 & ""

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
If Me.xTransCode1.Text <> "" And Me.xTransCode2.Text <> "" Then
    MsgBox " „  ”ÃÌ· ⁄„Ì· Ê „Ê—œ ·‰›” «·‘Ìﬂ"
    Exit Function
End If

If xCode1.Text <> "" Then
    If GetDesca("select code from file3_10 where code = " & MyParn(xCode1.Text)) = "" Then Exit Function
End If

If xCode2.Text <> "" Then
    If GetDesca("select code from file4_10 where code = " & MyParn(xCode2.Text)) = "" Then Exit Function
End If

If (xTransCode1.Text <> "" Or xTransCode2.Text <> "" Or Not xClosed(0).Value) And Not IsDate(xDATE_3.Text) Then
    MsgBox " ”ÃÌ·  «—ÌŒ «· ŸÂÌ— °  ÕœÌœ «‰ «·‘Ìﬂ  „  ŸÂÌ—…"
    Exit Function
End If

If XSER_NO.Enabled And Trim(xCode1.Text) <> "" Then
    cString = GetDesca("select ser_no from FILE5_20 where code1 = " & MyParn(xCode1.Text) & " and Chk_Id = " & MyParn(XCHK_ID.Text)) & ""
    If Trim(cString) <> "" Then
        MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” " & IIf(publicFlag = 1, " «·⁄„Ì· ", " «·„Ê—œ ")
        Exit Function
    End If
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
myDefine
On Error Resume Next
XCHK_ID.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
con.BeginTrans
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.Execute "delete  From FILE5_20 where Ser_No = " & MyParn(XSER_NO.Text)
End If
con.CommitTrans

CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "SER_NO < " & MyParn(XSER_NO.Text), , adSearchBackward, adBookmarkLast
    If CardTable.EOF Then CardTable.MoveFirst
    MyLoad
Else
    If optclose(0).Value Then CmdAdd_Click Else myDefine
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
Private Sub cmdSave_Click()
'msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
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
        MyLoad
    Else
        myDefine
    End If
End If
End Sub
Private Sub CmdUndo_Click()
bSumMode = False
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If XSER_NO.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "ser_no = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    MyLoad
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
lblDateBank.Visible = True
xDateBank.Visible = True
    
OpenCon con
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
Private Sub optclose_Click(Index As Integer)
cString = "Select FILE5_20.*,FILE3_10.DESCA AS DESCA1,FILE4_10.DESCA AS DESCA2" & _
          " From (FILE5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE) LEFT JOIN FILE4_10 ON FILE5_20.CODE2 = FILE4_10.CODE"
If Index <> 3 Then
    cString = cString & turn(cString) & " closed = " & MyParn(Index)
End If
cString = cString & " Order by Ser_No"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenKeyset, adLockOptimistic, adCmdText

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    If optclose(0).Value Then
        CmdAdd_Click
    Else
        myDefine
    End If
End If
End Sub
Private Sub XBOX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XBOX.BoundText = ""
End Sub

Private Sub xBOX2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xbox2.BoundText = ""
End Sub

Private Sub XCHK_ID_Validate(Cancel As Boolean)
If publicFlag = 2 And XSER_NO.Enabled Then
    If GetDesca("SELECT SER_NO FROM FILE5_20 WHERE SER_NO <> " & MyParn(XSER_NO.Text) & " AND  CHK_ID = " & MyParn(XCHK_ID.Text)) <> "" Then
        MsgBox "—ﬁ„ «·‘Ìﬂ „ﬂ—— „‰ ﬁ»·"
    End If
End If
End Sub
Private Sub xClosed_Click(Index As Integer)
xTransCode1.Enabled = (Index = 2)
xTransCode2.Enabled = (Index = 2)
If Index <> 2 Then
    xTransCode1.Text = ""
    xTransCode2.Text = ""
    xTransName1.Text = ""
    xTransName2.Text = ""
End If
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
xNAME1.Text = ""
If Trim(xCode1.Text) = "" Then Exit Sub
xCode1.Text = RetZero(xCode1.Text)
xNAME1.Text = GetDesca("Select Desca from FILE3_10 where code = " & MyParn(xCode1.Text))
End Sub
Private Sub xCode2_Change()
If xCode2.Text <> "" Then
    xCode1.Text = ""
    xCode1.Enabled = False
    xNAME1.Text = ""
Else
    xCode1.Enabled = True
End If
End Sub

Private Sub xCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub

Private Sub xCODE2_LostFocus()
xName2.Text = ""
xCode2.Text = RetZero(xCode2.Text)
xName2.Text = GetDesca("Select Desca from FILE4_10 where code = " & MyParn(xCode1.Text))
End Sub
Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""
End Sub
Private Sub xSer_no_LostFocus()
If Trim(XSER_NO.Text) = "" Then Exit Sub
XSER_NO.Text = RetZero(XSER_NO.Text, 5)
CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
Private Function MyReplace() As Boolean
Dim aInsert(22, 1)
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
aInsert(5, 1) = addstring(xdesca.Text)

aInsert(6, 0) = "Bank_rec"
aInsert(6, 1) = addstring(xBANK_REC.Text)

aInsert(7, 0) = "date_1"
aInsert(7, 1) = addDate(xDATE_1.Text)

aInsert(8, 0) = "date_3"
aInsert(8, 1) = addDate(xDATE_3.Text)

aInsert(9, 0) = "date_r"
aInsert(9, 1) = addDate(xDATE_R.Text)

aInsert(10, 0) = "DateBank"
aInsert(10, 1) = addDate(xDateBank.Text)

aInsert(11, 0) = "[VALUE]"
aInsert(11, 1) = Val(xValue.Text)

aInsert(12, 0) = "NAME4"
aInsert(12, 1) = addstring(xNAME4.Text)

aInsert(13, 0) = "BOX"
aInsert(13, 1) = addstring(XBOX.BoundText)

aInsert(14, 0) = "MEMO"
aInsert(14, 1) = addstring(xMEMO.Text)

aInsert(15, 0) = "ID_BANK"
aInsert(15, 1) = addstring(XID_BANK.BoundText)

If xCode1.Text <> "" Then
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = addstring("⁄„Ì·:" & xNAME1.Text)
Else
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = "NULL"
End If

If xCode2.Text <> "" Then
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = addstring("„Ê—œ:" & xName2.Text)
Else
    aInsert(16, 0) = "DESCA1"
    aInsert(16, 1) = "NULL"
End If

If xTransCode1.Text <> "" Then
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = addstring("„Ê—œ: " & xTransName1.Text)
Else
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = "NULL"
End If

If xTransCode2.Text <> "" Then
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = addstring("⁄„Ì·: " & xTransName1.Text)
Else
    aInsert(17, 0) = "descaTrans"
    aInsert(17, 1) = "NULL"
End If
 
aInsert(18, 0) = "Closed"
aInsert(18, 1) = addstring(retClose)

aInsert(19, 0) = "TransCode1"
aInsert(19, 1) = addstring(xTransCode1.Text)

aInsert(20, 0) = "TransCode2"
aInsert(20, 1) = addstring(xTransCode2.Text)

aInsert(21, 0) = "TransName1"
aInsert(21, 1) = addstring(xTransName1.Text)

aInsert(22, 0) = "TransName2"
aInsert(22, 1) = addstring(xTransName2.Text)

On Error GoTo myerror
con.BeginTrans
If XSER_NO.Enabled Then
    XSER_NO.Text = RetZero(Val(Newflag("file5_20", "ser_no")), 5)
    aInsert(0, 1) = addstring(XSER_NO.Text)
    con.Execute CreateInsert(aInsert, "file5_20")
Else
    con.Execute CreateUpdate(aInsert, "file5_20", " where SER_NO = " & addstring(XSER_NO.Text))
End If
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub handleopt()
For i = 1 To 3
    optclose(i).Enabled = GetDesca("select ser_no from " & cFileName & " where closed = " & MyParn(i)) <> ""
Next
End Sub
Private Sub ChqLookup()
Dim Generalarray(4)
Dim GrdArray(5)
Set Generalarray(1) = Me
If publicFlag = 1 Then
    Generalarray(2) = "Select Ser_No as [„”·”· «·‘Ìﬂ],Name1 As [⁄„Ì· ],TRANSNAME1 As [„Ê—œ ],FORMAT(Date_1,'DD-MM-YYYY') As [ «—ÌŒ «” Õﬁ«ﬁ «·‘Ìﬂ],Value as [ﬁÌ„… «·‘Ìﬂ] From " & cFileName & "  Where Closed = " & MyParn(sClose)
Else
    Generalarray(2) = "Select Ser_No as [„”·”· «·‘Ìﬂ],Name1 As [„Ê—œ ],Name12 As [⁄„Ì·],FORMAT(Date_1,'DD-MM-YYYY') As [ «—ÌŒ «” Õﬁ«ﬁ «·‘Ìﬂ],Value as [ﬁÌ„… «·‘Ìﬂ] From " & cFileName & "  Where Closed = " & MyParn(sClose)
End If
Generalarray(3) = " and ( Name1 Like '*cFilter*' Or NAME12 Like '*cFilter*') "
Generalarray(4) = "Order By Name1 , name12 "
   
GrdArray(1) = 1000
GrdArray(2) = 2000
GrdArray(3) = 2000
GrdArray(4) = 1200
GrdArray(5) = 1200

Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search3.Caption = "«” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub xTransCode1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP 2
End Sub
Private Sub xTransCode1_LostFocus()
If Trim(xTransCode1.Text) <> "" Then
    xTransName1.Text = GetDesca("Select Desca from file4_10 where code = " & xTransCode1.Text)
End If

If xTransCode1.Text <> "" Then
    xTransCode2.Text = ""
    xTransCode2.Enabled = False
    xTransName2.Text = ""
Else
    xTransCode2.Enabled = True
End If
End Sub

Private Sub xTransCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP
End Sub
Private Sub xTransCode2_LostFocus()
If Trim(xTransCode2.Text) <> "" Then
    xTransName2.Text = GetDesca("Select Desca from file3_10 where code = " & xTransCode2.Text)
    xTransCode1.Text = ""
    xTransCode1.Enabled = False
    xTransName1.Text = ""
Else
    xTransCode1.Enabled = True
End If
End Sub
Private Function retClose() As String
Dim i As Integer
For i = 0 To 2
    If xClosed(i).Value Then
        retClose = i & ""
        Exit For
    End If
Next
End Function
