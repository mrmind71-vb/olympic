VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form chq 
   Caption         =   "‘Ìﬂ« "
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
         Picture         =   "Chq.frx":0000
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
         Picture         =   "Chq.frx":0532
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
         Picture         =   "Chq.frx":067C
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
         Picture         =   "Chq.frx":0BAE
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
         Picture         =   "Chq.frx":0FF0
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
         Picture         =   "Chq.frx":1522
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
         Picture         =   "Chq.frx":1FF3
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
Attribute VB_Name = "chq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset, RecordCountTable As ADODB.Recordset
Dim cFileName As String, ChargeTable As ADODB.Recordset
Dim ClientTable1 As ADODB.Recordset, ClientTable2 As ADODB.Recordset
Dim cClient1 As String, cClient2 As String
Dim cFileMove1 As String, cFileMove2 As String, cChqDesc As String
Dim cFieldUnder As String, cFieldTrans, cFieldReject As String
Dim CMOVE As String, bSumMode As Boolean
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And optclose(0).Value) And bEdit
cmdSave.Enabled = (nMode = LoadMode Or optclose(0).Value) And bEdit
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
Dim GrdArray()

If publicFlag = 1 Then ReDim GrdArray(6, 1)
If publicFlag = 2 Then ReDim GrdArray(5, 1)

If optclose(0).Value Then cWhere = " WHERE CLOSED = '0'"
If optclose(1).Value Then cWhere = " WHERE CLOSED = '1'"
If optclose(2).Value Then cWhere = " WHERE CLOSED = '2'"

Set Generalarray(0) = Me
If publicFlag = 1 Then
    Generalarray(1) = " SELECT FILE5_20.SER_NO AS „”·”· ,  FILE5_20.DESCA1 as [‘Ìﬂ „‰], format(FILE5_20.DATE_1 ,'d-m-yyyy') AS ≈” Õﬁ«ﬁ, format(FILE5_20.DATE_r ,'d-m-yyyy') AS  Õ—Ì—, VALUE AS ﬁÌ„… , CHK_ID as [—ﬁ„ «·‘Ìﬂ], FILE5_20.TRANSNAME1 AS  ŸÂÌ—   From FILE5_20 " & cWhere
Else
    Generalarray(1) = " SELECT file5_21.SER_NO AS „”·”· ,  file5_21.DESCA1 as [‘Ìﬂ ·], format(file5_21.DATE_1 ,'d-m-yyyy') AS ≈” Õﬁ«ﬁ, format(file5_21.DATE_r ,'d-m-yyyy') AS  Õ—Ì—  , VALUE AS ﬁÌ„… , CHK_ID as [—ﬁ„ «·‘Ìﬂ] From file5_21 " & cWhere
End If

Generalarray(2) = ""
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "„”·”· «Ê „” ›Ìœ -  «—ÌŒ  Õ—Ì—-—ﬁ„ «·‘Ìﬂ"
listarray(0, 1) = "(%%desca1%% or chk_id Like '%cFilter%' Or ##Date_R##)"

listarray(1, 0) = " «—ÌŒ «” Õﬁ«ﬁ"
listarray(1, 1) = " ##Date_1##"
'listarray(1, 5) = 10

listarray(2, 0) = "«·ﬁÌ„…"
listarray(2, 1) = "##[value]##"

GrdArray(0, 0) = "„”·”·"
GrdArray(0, 1) = 800

GrdArray(1, 0) = IIf(publicFlag = 1, "‘Ìﬂ „‰", "‘Ìﬂ ·‹")
GrdArray(1, 1) = 2000

GrdArray(2, 0) = "«” Õﬁ«ﬁ"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = " Õ—Ì—"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "ﬁÌ„…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "—ﬁ„ «·‘Ìﬂ"
GrdArray(5, 1) = 1200
If publicFlag = 1 Then
    GrdArray(6, 0) = " ŸÂÌ—"
    GrdArray(6, 1) = 2000
End If
searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = cTitle
Search3.Show 1
End Sub
Sub CLIENTLOOKUP(pflag)
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT Code , Desca  From " & IIf(pflag = 1, cClient1, cClient2)
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
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
Sub myDefine()
XSER_NO.Text = RetZero(Val(Newflag(cFileName, "ser_no")), 5)
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
XSER_NO.Text = RetZero(Val(Newflag(cFileName, "ser_no")), 5)
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
XSER_NO.Text = CardTable!Ser_no & ""
XCHK_ID.Text = CardTable!CHK_ID & ""
xCode1.Text = CardTable!Code1 & ""
xCode2.Text = CardTable!Code2 & ""

xNAME1.Text = GetDesca("Select Desca from " & cClient1 & " where code = " & MyParn(xCode1.Text))
xName2.Text = GetDesca("Select Desca from " & cClient2 & " where code = " & MyParn(xCode2.Text))

xNAME4.Text = CardTable!NAME4 & ""

xdesca.Text = CardTable!Desca & ""
'xCharge.Text = TurnValue(CardTable!Charge, Null, "")

xClosed(0).Value = IIf(CardTable!CLOSED = "0", True, False)
xClosed(1).Value = IIf(CardTable!CLOSED = "1", True, False)
xClosed(2).Value = IIf(CardTable!CLOSED = "2", True, False)

xOld.Value = IIf(CardTable!OLD, 1, 0)
xBANK_REC.Text = CardTable!Bank_rec & ""
xValue.Text = Format(CardTable!Value, "Fixed")
xDATE_1.Text = Format(CardTable!date_1, "dd-mm-yyyy")
xDATE_3.Text = Format(CardTable!date_3, "dd-mm-yyyy")
xDATE_R.Text = Format(CardTable!date_R, "dd-mm-yyyy")
If publicFlag = 1 Then xDateBank.Text = Format(CardTable!DateBank, "dd-mm-yyyy")
xMEMO.Text = CardTable!Memo & ""
XID_BANK.BoundText = CardTable!ID_BANK & ""
XBOX.BoundText = CardTable!Box & ""
If publicFlag = 1 Then
    xTransCode1.Text = CardTable!TRANSCODE1 & ""
    xTransCode2.Text = CardTable!TRANSCODE2 & ""
    
    xTransName1.Text = CardTable!TRANSNAME1 & ""
    xTransName2.Text = CardTable!transname2 & ""
End If
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
'If xNAME1.Text = "" And xName2.Text = "" And xDesca.Text = "" Then
'    MsgBox "«·≈”„  ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
'    Exit Function
'End If
If Me.xTransCode1.Text <> "" And Me.xTransCode2.Text <> "" Then
    MsgBox " „  ”ÃÌ· ⁄„Ì· Ê „Ê—œ ·‰›” «·‘Ìﬂ"
    Exit Function
End If

If xCode1.Text <> "" Then
    If GetDesca("select code from " & cClient1 & " where code = " & MyParn(xCode1.Text)) = "" Then Exit Function
End If

If xCode2.Text <> "" Then
    If GetDesca("select code from " & cClient2 & " where code = " & MyParn(xCode2.Text)) = "" Then Exit Function
End If

If (xTransCode1.Text <> "" Or xTransCode2.Text <> "" Or Not xClosed(0).Value) And Not IsDate(xDATE_3.Text) Then
    MsgBox " ”ÃÌ·  «—ÌŒ «· ŸÂÌ— °  ÕœÌœ «‰ «·‘Ìﬂ  „  ŸÂÌ—…"
    Exit Function
End If
If publicFlag = 1 Then
    If XSER_NO.Enabled And Trim(xCode1.Text) <> "" Then
        cString = GetDesca("select ser_no from " & cFileName & " where code1 = " & MyParn(xCode1.Text) & " and Chk_Id = " & MyParn(XCHK_ID.Text)) & ""
        If Trim(cString) <> "" Then
            MsgBox "‘Ìﬂ »‰›” «·—ﬁ„ ·‰›” " & IIf(publicFlag = 1, " «·⁄„Ì· ", " «·„Ê—œ ")
            Exit Function
        End If
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
    con.Execute "delete  From " & cFileName & " where Ser_No = " & MyParn(XSER_NO.Text)
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
If publicFlag = 1 Then
    Me.Caption = "√Ê—«ﬁ ﬁ»÷"
    cChqDesc = "√Ê—«ﬁ ﬁ»÷"
    cFileName = "File5_20"
    cClient1 = "File3_10"
    cClient2 = "File4_10"
    cFileMove1 = "File3_11"
    cFileMove2 = "File4_11"
    lblDateBank.Visible = True
    xDateBank.Visible = True
Else
    Me.Caption = "√Ê—«ﬁ œ›⁄"
    cChqDesc = "√Ê—«ﬁ œ›⁄"
    cFileName = "File5_21"
    cClient1 = "File4_10"
    cClient2 = "File3_10"
    cFileMove1 = "File4_11"
    cFileMove2 = "File3_11"
    cDelMove1 = " "
    cDelMove2 = " "
    LabelCode.Caption = "ﬂÊœ „Ê—œ"
    LabelCode2.Caption = "ﬂÊœ ⁄„Ì·"
    frmTrans.Visible = False
    Label2.Caption = "»‰ﬂ «·”Õ»"
    xdesca.Height = xdesca.Height - xDateBank.Height
    Frame1.Top = Frame1.Top - xDateBank.Height
    Me.Height = Me.Height - frmTrans.Height - 200
End If

Set alltable = New ADODB.Recordset
alltable.Open "Select * from " & cFileName & "  order by Ser_No", con, adOpenStatic, adLockReadOnly, adCmdText

Set ClientTable1 = New ADODB.Recordset
ClientTable1.Open cClient1, con, adOpenStatic, adLockReadOnly, adCmdTable

Set ClientTable2 = New ADODB.Recordset
ClientTable2.Open cClient2, con, adOpenStatic, adLockReadOnly, adCmdTable


data1.ConnectionString = con.ConnectionString
data1.RecordSource = "FILE0_50"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

DATA2.ConnectionString = con.ConnectionString
DATA2.RecordSource = "FILE5_10"

Set XID_BANK.RowSource = DATA2
XID_BANK.ListField = "Desca"
XID_BANK.BoundColumn = "code"

optclose(0).Value = True
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    CmdAdd_Click
End If
End Sub

Private Sub optclose_Click(Index As Integer)
If Index <> 3 Then
    cString = "Select * From " & cFileName & " Where closed = " & MyParn(Index) & " Order by Ser_No "
Else
    cString = "Select * From " & cFileName & " Where TRUE Order by Ser_No "
End If

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
    If GetDesca("SELECT SER_NO FROM FILE5_21 WHERE SER_NO <> " & MyParn(XSER_NO.Text) & " AND  CHK_ID = " & MyParn(XCHK_ID.Text)) <> "" Then
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
Private Sub xCODE1_DblClick()
CLIENTLOOKUP 1
End Sub
Private Sub xCode1_LostFocus()
xNAME1.Text = ""
If Trim(xCode1.Text) = "" Then Exit Sub
xCode1.Text = RetZero(xCode1.Text)
'ClientTable1.Find "code = " & MyParn(xCode1.Text), , adSearchForward, adBookmarkFirst
'If Not ClientTable1.EOF Then
'    xNAME1.Text = ClientTable1!desca
'End If
xNAME1.Text = GetDesca("Select Desca from " & cClient1 & " where code = " & MyParn(xCode1.Text))
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
Private Sub xCODE2_DblClick()
CLIENTLOOKUP 2
End Sub
Private Sub xCODE2_LostFocus()
xName2.Text = ""
xCode2.Text = RetZero(xCode2.Text)
'ClientTable2.Find "code = " & MyParn(xCode2.Text), , adSearchForward, adBookmarkFirst
'If Not ClientTable2.EOF Then
'    xName2.Text = ClientTable2!desca
'End If
xName2.Text = GetDesca("Select Desca from " & cClient2 & " where code = " & MyParn(xCode1.Text))
End Sub
Private Sub xCode3_Change()
If xCode3.Text <> "" Then
    xCode1.Text = ""
    xCode1.Enabled = False
    xNAME1.Text = ""
    xCode2.Text = ""
    xCode2.Enabled = False
    xName2.Text = ""
Else
    xCode1.Enabled = True
    xCode2.Enabled = True
End If
End Sub
Private Sub xCode3_DblClick()
    Dim Generalarray(3)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code as «·ﬂÊœ,DescA as [«·„’—Ê›] From file8_70"
    Generalarray(3) = " Where DescA Like '%cFilter%'"
           
    GrdArray(1) = 1200
    GrdArray(2) = 4000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search3.Caption = "«” ⁄·«„ "
    Search3.Show 1
End Sub
Private Sub xCode3_LostFocus()
xName3.Text = ""
ChargeTable.Find "code = " & MyParn(xCode3.Text), , adSearchForward, adBookmarkFirst
If Not ChargeTable.EOF Then
    xName3.Text = ChargeTable!Desca
End If
End Sub
Private Sub XID_BANK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then XID_BANK.BoundText = ""
End Sub

Private Sub xSer_no_LostFocus()
XSER_NO.Text = RetZero(XSER_NO.Text, 5)
CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
Sub MyReplace()
CardTable.Find "SER_NO = " & MyParn(XSER_NO.Text), , adSearchForward, adBookmarkFirst
'On Error GoTo Myerror
con.BeginTrans
If CardTable.EOF Then CardTable.AddNew
CardTable!Ser_no = TurnValue(XSER_NO.Text, "", Null)
CardTable!CHK_ID = TurnValue(XCHK_ID.Text, "", Null)
CardTable!OLD = IIf(xOld.Value = 0, False, True)
CardTable!Code1 = TurnValue(xCode1.Text, "", Null)
CardTable!Code2 = TurnValue(xCode2.Text, "", Null)
CardTable!Desca = TurnValue(xdesca.Text, "", Null)
CardTable!Bank_rec = TurnValue(xBANK_REC.Text, "", Null)
CardTable!date_1 = DateFix(xDATE_1.Text)
CardTable!date_3 = DateFix(xDATE_3.Text)
CardTable!date_R = DateFix(xDATE_R.Text)
If publicFlag = 1 Then CardTable!DateBank = DateFix(xDateBank.Text)

'CardTable!TRANSCODE1 = TurnValue(xTransCode1.Text, "", Null)
'CardTable!TRANSCODE2 = TurnValue(xTransCode2.Text, "", Null)

CardTable!Value = Val(xValue.Text)
CardTable!NAME4 = TurnValue(xNAME4.Text, "", Null)
CardTable!Box = TurnValue(XBOX.BoundText, "", Null)
CardTable!Memo = TurnValue(xMEMO.Text, "", Null)
CardTable!ID_BANK = TurnValue(XID_BANK.BoundText, "", Null)
cString = ""
If xCode1.Text <> "" Then
    cString = IIf(publicFlag = 1, "⁄„Ì·:", "„Ê—œ:") & xNAME1.Text
End If

If xCode2.Text <> "" Then
   cString = IIf(publicFlag <> 1, "⁄„Ì·:", "„Ê—œ:") & xNAME1.Text
End If

CardTable!desca1 = TurnValue(cString, "", Null)

cString = ""
If publicFlag = 1 Then
    If xTransCode1.Text <> "" Then
        cString = "„Ê—œ:" & xTransName1.Text
    End If
    
    If xTransCode2.Text <> "" Then
        cString = "⁄„Ì·:" & xTransName2.Text
    End If
    CardTable!descatrans = TurnValue(cString, "", Null)
End If
If xClosed(0).Value Then
    CardTable!CLOSED = 0
End If
If xClosed(1).Value Then
    CardTable!CLOSED = 1
End If
If xClosed(2).Value Then
    CardTable!CLOSED = 2
End If
If publicFlag = 1 Then
    CardTable!TRANSCODE1 = TurnValue(xTransCode1.Text, "", Null)
    CardTable!TRANSCODE2 = TurnValue(xTransCode2.Text, "", Null)
    CardTable!TRANSNAME1 = TurnValue(xTransName1.Text, "", Null)
    CardTable!transname2 = TurnValue(xTransName2.Text, "", Null)
End If
CardTable.Update
con.CommitTrans
CardTable.Requery
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Cmd_Trans_Click()
If IsDate(xDATE_3.Text) Then
    cmdSave_Click
    If XSER_NO.Text <> "" And xTransCode1.Text <> "" Then
        con.Execute " UPDATE FILE5_20 SET FILE5_20.CLOSED  = '2' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        con.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
        cString = "Insert Into File4_11(" & _
                  "[Type],Doc_Id,Code,[Date],PAY,DescA)" & _
                  " Select 'T',Ser_No, TRANSCODE1 ,Date_3,Value,'  ŸÂÌ— ‘Ìﬂ Õﬁ ' & date_1" & _
                  " From File5_20" & _
                  " WHERE FILE5_20.SER_NO  = " & MyParn(XSER_NO.Text)
        con.Execute cString
    End If
    
    If xCode2.Text <> "" And xTransCode2.Text <> "" Then
        con.Execute " UPDATE FILE5_20 SET FILE5_20.CLOSED  = '2' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        cmdSave_Click
        con.Execute "Delete * From File3_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
        cString = "Insert Into File3_11(" & _
                  "[Type],Doc_Id,Code,[Date],SAL,DescA)" & _
                  " Select 'T',Ser_No,TRANSCODE2 ,Date_3,Value,'  ŸÂÌ— ‘Ìﬂ Õﬁ ' & date_1" & _
                  " From File5_20" & _
                  " WHERE FILE5_20.SER_NO  = " & MyParn(XSER_NO.Text)
        con.Execute cString
    End If
Else
    MsgBox " ”ÃÌ·  «—ÌŒ  ŸÂÌ— «·‘Ìﬂ "
End If
End Sub
Private Sub Cmd_unTrans_Click()
If IsDate(xDATE_3.Text) Then
    xTransCode1.Text = ""
    xTransCode2.Text = ""
    cmdSave_Click
    If xCode.Text <> "" And xTransCode1.Text <> "" Then
        con.Execute " UPDATE FILE5_20 SET FILE5_20.CLOSED  = '0' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        con.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
    End If
    
    If xCode2.Text <> "" And xTransCode2.Text <> "" Then
        con.Execute " UPDATE FILE5_20 SET FILE5_20.CLOSED  = '0' WHERE SER_NO = " & MyParn(XSER_NO.Text)
        con.Execute "Delete * From File4_11 Where Doc_ID = " & MyParn(XSER_NO.Text) & " and [type] = 'T'"
    End If
End If
End Sub
Private Sub CmdColect_Click()
XBOX.Enabled = True
xDATE_3.Enabled = True
xMEMO.Enabled = True
xDATE_3.SetFocus
End Sub
Private Sub CmdDelColect_Click()
If MsgBox("Â·  Êœ«·€«¡  Õ’Ì· «·‘Ìﬂ", vbOKCancel, "«·„—‘œ") = vbCancel Then Exit Sub
xDATE_3.Text = ""
xMEMO.Text = ""
CardTable!date_3 = Null
CardTable!CLOSED = "0"
CardTable!Memo = Null
CardTable.Update
CardTable.Requery

If CardTable.RecordCount = 0 Then
    optclose(2).Enabled = False
    optclose(0).Value = True
Else
    CardTable.MoveLast
    MyLoad
End If
End Sub
Private Sub MYREPLACEMOVE()
        
    cField = Me.Caption & " " & "Õﬁ" & Format(xDATE_1.Text, "d-m-yyyy")
    cFieldR = Me.Caption & " " & " „— œ Õﬁ " & Format(xDATE_1.Text, "d-m-yyyy")
   
    ' ‘Ìﬂ  Õ  «· Õ’Ì·
    If xOld.Value = 0 Then
        If xCode1.Text <> "" Then
            cString = "Insert Into " & cFileMove1 & _
                      "([Type],Doc_Id,Code,[Date],PAY,CHQ,DescA) " & _
                      " VALUES ( " & _
                      addstring("A") & "," & _
                      addstring(XSER_NO.Text) & "," & _
                      addstring(xCode1.Text) & "," & _
                      DateSql(xDATE_R.Text) & "," & _
                      addvalue(xValue.Text) & "," & _
                      addstring(XSER_NO.Text) & "," & _
                      addstring(cField) & _
                      ")"
            con.Execute cString
        End If
        
        If xCode2.Text <> "" Then
            cString = "Insert Into " & cFileMove2 & " (" & _
                      "[Type],Doc_Id,Code,[Date],SAL,DescA) " & _
                      " Values ( " & _
                      addstring("B") & "," & _
                      addstring(XSER_NO.Text) & "," & _
                      addstring(xCode2.Text) & "," & _
                      DateSql(xDATE_R.Text) & "," & _
                      addvalue(xValue.Text) & "," & _
                      addstring(cField) & _
                      ")"
            con.Execute cString
        End If
    End If
    
    
   ' —›÷ ‘Ìﬂ
   If xClosed(1).Value Then
        If xCode1.Text <> "" Then
            cString = "Insert Into " & cFileMove1 & _
                      "([Type],Doc_Id,Code,[Date],SAL,DescA) " & _
                      " VALUES ( " & _
                      addstring("C") & "," & _
                      addstring(XSER_NO.Text) & "," & _
                      addstring(xCode1.Text) & "," & _
                      DateSql(xDATE_R.Text) & "," & _
                      addvalue(xValue.Text) & "," & _
                      addstring(cFieldR) & _
                      ")"
            con.Execute cString
        End If
        
        If xCode2.Text <> "" Then
            cString = "Insert Into " & cFileMove2 & _
                      "([Type],Doc_Id,Code,[Date],PAY,DescA) " & _
                     " VALUES ( " & _
                      addstring("D") & "," & _
                      addstring(XSER_NO.Text) & "," & _
                      addstring(xCode2.Text) & "," & _
                      DateSql(xDATE_R.Text) & "," & _
                      addvalue(xValue.Text) & "," & _
                      addstring(cFieldR) & _
                      ")"
                      con.Execute cString
        End If
   End If
 
 '  ŸÂÌ— ‘Ìﬂ
    If xTransCode1.Text <> "" Then
        cString = "Insert Into  FILE4_11 (" & _
                  "[Type],Doc_Id,Code,[Date],PAY,DescA) " & _
                  " VALUES ( " & _
                  addstring("E") & "," & _
                  addstring(XSER_NO.Text) & "," & _
                  addstring(xTransCode1.Text) & "," & _
                  DateSql(xDATE_3.Text) & "," & _
                  addvalue(xValue.Text) & "," & _
                  addstring("  ŸÂÌ— ‘Ìﬂ Õﬁ " & Format(date_1, "d-m-yyyy")) & _
                  ")"
        con.Execute cString
    End If
        
    If xTransCode2.Text <> "" Then
        cString = "Insert Into  FILE3_11 (" & _
                  "[Type],Doc_Id,Code,[Date],SAL,DescA) " & _
                  " VALUES ( " & _
                  addstring("F") & "," & _
                  addstring(XSER_NO.Text) & "," & _
                  addstring(xTransCode2.Text) & "," & _
                  DateSql(xDATE_3.Text) & "," & _
                  addvalue(xValue.Text) & "," & _
                  addstring("  ŸÂÌ— ‘Ìﬂ Õﬁ " & date_1) & _
                  ")"
        con.Execute cString
    End If
End Sub
Private Sub xSer_no__DblClick()
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As «·«”„Ê  From " & cFile_10
    Generalarray(3) = "Where DescA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 2600
    GrdArray(3) = 1500
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search3.Caption = "«” ⁄·«„ "
    Search3.Show 1
End Sub
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
Private Sub xTRANSCode1_Change()
ClientTable2.Find " code = " & MyParn(xTransCode1.Text), , adSearchForward, adBookmarkFirst
If Not ClientTable2.EOF Then
    xTransName1.Text = ClientTable2!Desca
Else
    xTransName1.Text = ""
End If

If xTransCode1.Text <> "" Then
    xTransCode2.Text = ""
    xTransCode2.Enabled = False
    xTransName2.Text = ""
Else
    xTransCode2.Enabled = True
End If
End Sub
Private Sub xTRANSCODE1_DblClick()
CLIENTLOOKUP 2
End Sub
Private Sub xTRANSCode2_Change()
ClientTable1.Find " code = " & MyParn(xTransCode2.Text), , adSearchForward, adBookmarkFirst
If Not ClientTable1.EOF Then
    xTransName2.Text = ClientTable1!Desca
Else
    xTransName2.Text = ""
End If
If xTransCode2.Text <> "" Then
    xTransCode1.Text = ""
    xTransCode1.Enabled = False
    xTransName1.Text = ""
Else
    xTransCode1.Enabled = True
End If
End Sub
Private Sub xTRANSCODE2_DblClick()
CLIENTLOOKUP 1
End Sub
Public Sub MoveLOadChq(CSER_NO)
    CardTable.Find "SER_NO = " & MyParn(CSER_NO), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        MyLoad
    Else
        myDefine
    End If
End Sub

