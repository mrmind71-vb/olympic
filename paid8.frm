VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form paid 
   Caption         =   "«·”œ«œ"
   ClientHeight    =   8220
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   13125
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2014-12-31"
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   -990
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   7515
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Frame Frame11 
      Height          =   1185
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   45
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton cmdcomp3 
         Caption         =   "«·ﬂ·"
         Enabled         =   0   'False
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cmdComp2 
         Caption         =   "€Ì— „”œœ…"
         Enabled         =   0   'False
         Height          =   420
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   630
         Width           =   1050
      End
      Begin VB.CommandButton cmdComp 
         Caption         =   "⁄„· «” „«—…"
         Enabled         =   0   'False
         Height          =   420
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   630
         Width           =   1410
      End
      Begin MSDBCtls.DBCombo xCompany 
         Bindings        =   "paid8.frx":0000
         Height          =   390
         Left            =   45
         TabIndex        =   48
         Top             =   180
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1260
      Width           =   1815
      Begin VB.CommandButton Command1 
         Caption         =   "„ÿ«·»… »œÊ‰  «»⁄Ì‰"
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1710
      End
      Begin VB.CommandButton Command2 
         Caption         =   " €ÌÌ— «Ì«„ «·€—«„…"
         Height          =   465
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   1695
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Height          =   945
      Left            =   315
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   1667
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   90
      Width           =   6405
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "«” ⁄·«„"
         Height          =   420
         Left            =   5220
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
         Height          =   420
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Õ–›"
         Height          =   420
         Left            =   1050
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "„ÿ«·»… ÃœÌœ…"
         Height          =   420
         Left            =   4065
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   420
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1260
      Width           =   1950
      Begin VB.CommandButton CmdPrint 
         Caption         =   "ÿ»«⁄… «·„ÿ«·»…"
         Height          =   435
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1845
      End
      Begin VB.CommandButton CmdDetial 
         Caption         =   "≈÷«›… »Ì«‰«  «·„ÿ«·»…"
         Height          =   480
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «”«”Ì…"
      Height          =   1725
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   9150
      Begin VB.CheckBox xStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„ Êﬁ›"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   270
         Width           =   1770
      End
      Begin MSDBCtls.DBCombo xType 
         Bindings        =   "paid8.frx":0014
         Height          =   315
         Left            =   3510
         TabIndex        =   44
         Top             =   945
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DBCombo1"
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
      Begin VB.TextBox xYear 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   150
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   975
         Width           =   990
      End
      Begin VB.TextBox xNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   150
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "2014-30-06"
         Top             =   1305
         Width           =   1320
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6480
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6480
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label xDebit_String 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1305
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label Label6 
         Caption         =   "‰Ê⁄ «·„ÿ«·»… :"
         Height          =   240
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label LblYear 
         Caption         =   "·„Ê”„ :"
         Height          =   240
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   990
         Width           =   765
      End
      Begin VB.Label xSectionName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label Label4 
         Caption         =   " «—ÌŒ «·”œ«œ :"
         Height          =   240
         Left            =   7875
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   1035
      End
      Begin VB.Label xLastYear 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
         Width           =   2940
      End
      Begin VB.Label xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label Label3 
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ… :"
         Height          =   240
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ﬁ”Ì„… :"
         Height          =   240
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "„”·”· :"
         Height          =   240
         Left            =   7905
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7065
      Width           =   2940
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label xTotal0 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "«·≈Ã„«·Ì :"
         Height          =   240
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "«·«‘ —«ﬂ «·”‰ÊÌ :"
         Height          =   240
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1410
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4305
      Top             =   8025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10215
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7065
      Width           =   2895
      Begin VB.TextBox xOverDue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   585
         Width           =   1005
      End
      Begin VB.TextBox xLate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "€—«„…  √ŒÌ— :"
         Height          =   240
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "≈‘ —«ﬂ«  „ √Œ—… :"
         Height          =   240
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   630
         Width           =   1425
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7110
      Width           =   2190
      Begin VB.CommandButton CmdLast 
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   75
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdFirst 
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   1125
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
         Width           =   990
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   75
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   150
         Width           =   990
      End
      Begin VB.CommandButton CmdPrevious 
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   1125
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   150
         Width           =   990
      End
   End
   Begin VB.Frame Frame8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   7125
      Width           =   1365
      Begin VB.CommandButton cmdLoadMem 
         Caption         =   "»Ì«‰«  «·«⁄÷«¡"
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   585
         Width           =   1275
      End
      Begin VB.CommandButton cmdpaidFix 
         Caption         =   "÷»ÿ «·„ÿ«·»…"
         Enabled         =   0   'False
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   135
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4560
      Left            =   0
      TabIndex        =   55
      Top             =   2385
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   8043
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "paid8.frx":0028
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "paid8.frx":0044
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "paid8.frx":0060
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "paid8.frx":007C
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Grid1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4020
         Index           =   0
         Left            =   -74910
         TabIndex        =   56
         Top             =   360
         Width           =   12795
         _cx             =   22569
         _cy             =   7091
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4020
         Index           =   1
         Left            =   -74910
         TabIndex        =   57
         Top             =   360
         Width           =   12795
         _cx             =   22569
         _cy             =   7091
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4020
         Index           =   2
         Left            =   -74910
         TabIndex        =   58
         Top             =   360
         Width           =   12795
         _cx             =   22569
         _cy             =   7091
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4020
         Index           =   3
         Left            =   90
         TabIndex        =   59
         Top             =   360
         Width           =   12795
         _cx             =   22569
         _cy             =   7091
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
End
Attribute VB_Name = "paid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bNoRelation As Boolean, sSystem As String
Dim bNoSave As Boolean, bAdd As Boolean
Dim bNoChangeDate As Boolean
Dim nMulti As Integer
Dim oSearchComp As New Search3, oSearchMember As New Search3
Dim nNewMeeting As Double
Dim SectionTable As Recordset
Dim ItemTable As Recordset, itemSectionTable As Recordset
Dim MemItemTable As Recordset
Dim paidTypeTable As Recordset
Dim MeetingTable As Recordset
Dim tPayItem As Recordset
Dim CardTable As Recordset
Dim temptable As Recordset
Dim formMode As Byte
Const LoadMode = 0, DefineMode = 1
Private Function MyReplace() As Boolean
On Error Resume Next
Do
    Err.Clear
    If xDoc_No.Enabled Or bAdd Then
        CardTable.AddNew
    Else
        CardTable.Edit
    End If
    CardTable!doc_no = xDoc_No.Text
    CardTable!CODE = xCode.Text
    CardTable!NO = TurnValue(xNo.Text, "", Null)
    CardTable!Date = DateFix(xDate.Text)
    CardTable!TOTAL = Val(xTotal.Caption)
    CardTable!Total0 = TurnValue(Val(xTotal0.Caption), 0, Null)
    CardTable!OverDue = TurnValue(Val(xOverDue.Text), 0, Null)
    CardTable!LATE = TurnValue(Val(xLate.Text), 0, Null)
    CardTable!Type = xType.BoundText
    CardTable!Year = Val(xYear.Text)
    CardTable!AddOverDue = (Val(xOverDue.Text) <> 0 And CalcOverDue = 0)
    CardTable.Update
    
    ' Ã“¡ «÷«›Ì · ÿ»Ìﬁ«  «·‘»ﬂ…
    If Err.Number = 3022 Then
        xDoc_No.Text = RetZero(Val(xDoc_No.Text) + 1)
    ElseIf Err.Number = 3020 Then
        xCode.Enabled = True
        If MsgBox(" „ Õ–› «·”Ã· „‰ ﬁ»· „” Œœ„ «Œ— !! «÷«›… «·”Ã·ø ", vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
        End If
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Function
    End If
Loop Until Err.Number = 0

' ============
Me.MousePointer = 11

mydb.Execute "Delete * from File2_30 where Doc_No = " & MyParn(xDoc_No.Text)

For nTab = 0 To 3
    For I = 1 To Grid1(nTab).rows - 2
        mydb.Execute "Insert Into File2_30(Doc_no,item,[Value],[Number],Discount,Rate,total,[Note],[year]) values (" & _
                      addstring(xDoc_No.Text) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 0)) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 2)) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 3)) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 4)) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 5)) & "," & _
                      addvalue(Grid1(nTab).TextMatrix(I, 6)) & "," & _
                      addstring(Grid1(nTab).TextMatrix(I, 7)) & "," & _
                      addvalue(nTab) & ")"
    Next
Next

FixMemberPaid (xCode.Text)
Me.MousePointer = 0
MyReplace = True
End Function
Sub myProc()
    Select Case ActiveControl.Name
    Case xDoc_No.Name
        xDoc_No.Text = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
        Unload Search3
    Case xCode.Name
        If Trim(oSearchMember.Grid1.TextMatrix(oSearchMember.Grid1.Row, 2)) <> "" Then
            MsgBox "«·⁄„Ì· „ Êﬁ›"
            Exit Sub
        End If
        ActiveControl.Text = oSearchMember.Grid1.TextMatrix(oSearchMember.Grid1.Row, 0)
        xCode_LostFocus
        Unload oSearchMember
    Case CmdInform.Name
        xDoc_No.Text = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
        CardTable.Seek "=", xDoc_No.Text
        If Not CardTable.NoMatch Then MyLoad
        Unload Search2
    Case "Grid1"
        With ActiveControl
         If .Row = .rows - 1 Then .AddItem ""
        .TextMatrix(.Row, 0) = Search.Grid1.TextMatrix(Search.Grid1.Row, 0)
        .TextMatrix(.Row, 1) = Search.Grid1.TextMatrix(Search.Grid1.Row, 1)
        .TextMatrix(.Row, 2) = RetValue(Search.Grid1.TextMatrix(Search.Grid1.Row, 0), ActiveControl.index)
        .TextMatrix(.Row, 4) = RetValue(Search.Grid1.TextMatrix(Search.Grid1.Row, 0), ActiveControl.index, 1)
         'GrdRow ActiveControl.Index
         Calctotals
         End With
    Case cmdComp.Name
        CmdAdd_Click
        xCode.Text = oSearchComp.Grid1.TextMatrix(oSearchComp.Grid1.Row, 0)
        xCode_LostFocus
        CmdDetial_Click
        
        If Not bNoSave Then
            doprint 1
            bAdd = True
            CmdSave_Click
            bAdd = False
            Unload oSearchComp
            oSearchComp.Show 1
        Else
            bNoSave = False
            Unload oSearchComp
            Member.cCode = xCode.Text
            Member.Show 1
        End If
    Case cmdComp2.Name, cmdcomp3.Name
        xDoc_No.Text = oSearchComp.Grid1.TextMatrix(oSearchComp.Grid1.Row, 0)
        CardTable.Seek "=", xDoc_No.Text
        If Not CardTable.NoMatch Then MyLoad
        Unload oSearchComp
End Select
End Sub

Private Sub cmdComp_Click()
CompLookup
End Sub

Private Sub cmdComp2_Click()
CompLookup2 "ISNULL([NO])"
End Sub

Private Sub cmdcomp3_Click()
CompLookup2
End Sub

Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”œ«œ : Â· «‰  „Ê«›ﬁ ø", 4, systemName) = 6 Then
    mydb.Execute " DELETE * FROM FILE2_20 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute " DELETE * FROM  FILE2_30 WHERE DOC_NO = " & MyParn(xDoc_No.Text)
    
    FixMemberPaid (xCode.Text)
    If CardTable.RecordCount = 0 Then
        mydefine
    Else
        CardTable.Seek "<", xDoc_No.Text
        If CardTable.NoMatch Then CardTable.MoveFirst
        MyLoad
    End If
End If
End Sub
Private Sub CmdExit_Click()
  Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(3, 1)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select top 1000 File2_20.Doc_No,File1_10.Code,file1_10.DescA, [No],Format([Date],'yyyy/m/d') From (File2_20 Inner Join File1_10 On  File2_20.Code =  File1_10.Code) inner Join file0_50 on file2_20.[TYPE] = FILE0_50.CODE "
Generalarray(2) = "Order by File2_20.Date"
Generalarray(3) = 6000
Generalarray(4) = True
Generalarray(5) = True

listarray(0, 0) = "«·—ﬁ„ «Ê «·«”„  √Ê ‰Ê⁄ «·„ÿ«·»…"
listarray(0, 1) = "(FILE1_10.Code Like val('cFilter') or FILE1_10.DescA Like '%cFilter%' OR FILE0_50.DESCA LIKE '%cFilter%')"

listarray(1, 0) = "—ﬁ„ «·ﬁ”Ì„… «Ê —ﬁ„ «·„” ‰œ"
listarray(1, 1) = "(IIF(ISNULL([NO]),FALSE,VAL([NO]) = VAL('cFilter')) OR Val(File2_20.Doc_No) Like Val('cFilter'))"

listarray(2, 0) = "«· «—ÌŒ √Ê »«·„Ê”„"
listarray(2, 1) = "( iif(isDate('cFilter'),Format(Date,'dd-mm-yy') = Format('cFilter','dd-mm-yy'),false) OR IIF(ISNULL(FILE2_20.[YEAR]),FALSE,VAL(FILE2_20.[YEAR]) = VAL('cFilter')))"
listarray(3, 0) = "»œÊ‰ —ﬁ„ «” „«—… ‰⁄„-·«"
listarray(3, 1) = "(" & _
                  "('cFilter' = '‰⁄„' and isnull([NO])) or " & _
                  "('cFilter' = '·«' and Not isnull([NO]))" & _
                  ")"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·⁄÷Ê"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "≈”„ «·⁄÷Ê"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "—ﬁ„ «·«” „«—…"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ ⁄‰ „ÿ«·»…"
Search3.Show 1
End Sub
Private Sub CardLookup()
Dim Generalarray(4)
Dim listarray(5, 9)
Dim GrdArray(3, 2)

Set Generalarray(0) = Me
Generalarray(1) = "Select File2_20.Doc_No,DescA, [No],Format([Date],'yyyy/m/d') From File2_20 Inner Join File1_10 On  File2_20.Code =  File1_10.Code "
Generalarray(2) = "Order by File2_20.Date"
Generalarray(3) = 8100


listarray(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
listarray(0, 1) = "FILE1_10.Code Like cFilter"
listarray(0, 2) = ""

listarray(1, 0) = "«”„ «·⁄÷Ê"
listarray(1, 1) = "DescA Like '%cFilter%'"
listarray(1, 2) = ""
listarray(1, 3) = ""
listarray(1, 4) = "»ÕÀ »≈”„ «·⁄÷Ê"
listarray(1, 5) = ""


listarray(2, 0) = "—ﬁ„ «·ﬁ”Ì„…"
listarray(2, 1) = "IIF(ISNULL([NO]),FALSE,VAL([NO]) = VAL(cFilter))"
listarray(2, 2) = ""
listarray(2, 3) = ""
listarray(2, 4) = " »ÕÀ »—ﬁ„ «·ﬁ”Ì„…"
listarray(2, 5) = ""

listarray(3, 0) = "«· «—ÌŒ"
listarray(3, 1) = "iif(isDate('cFilter'),[date] = DateValue('cFilter'),False)"
listarray(3, 2) = ""
listarray(3, 3) = ""
listarray(3, 4) = "»ÕÀ »«· «—ÌŒ"

listarray(4, 0) = "«·”‰…"
listarray(4, 1) = "Year([date]) = cFilter"
listarray(4, 2) = ""
listarray(4, 3) = ""
listarray(4, 4) = "»ÕÀ »«·”‰…"

listarray(5, 0) = "—ﬁ„ «·„” ‰œ"
listarray(5, 1) = "File2_20.Doc_No Like '%cFilter%'"
listarray(5, 2) = ""
listarray(5, 3) = ""
listarray(4, 4) = "»ÕÀ »«·”‰…"

listarray(5, 0) = "‰Ê⁄ «·„ÿ«·»…"
listarray(5, 1) = "fILE2_20.Type Like '%cFilter%'"
listarray(5, 2) = ""
listarray(5, 3) = ""
listarray(5, 4) = "»ÕÀ »«·‰Ê⁄"
listarray(5, 5) = ""
listarray(5, 6) = ""
listarray(5, 7) = "Select * From File0_00 where Flag = 11 "
listarray(5, 8) = "Code"
listarray(5, 9) = "Desca"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 2) = 1000

GrdArray(1, 0) = "≈”„ «·⁄÷Ê"
GrdArray(1, 2) = 3000

GrdArray(2, 0) = "—ﬁ„ «·«” „«—…"
GrdArray(2, 2) = 1000

GrdArray(3, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(3, 2) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search2
Search2.Caption = "«” ⁄·«„ ⁄‰ „ÿ«·»…"
Search2.Show 1
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
Private Sub cmdpaidFix_Click()
'FixPaid
Fixmembers
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub CmdAdd_Click()
If CardTable.RecordCount > 0 Then
    xDoc_No.Text = IncRec(myLastField(CardTable, "DOC_NO"))
Else
    xDoc_No.Text = "000001"
End If
mydefine
End Sub
Private Sub CmdSave_Click()
If Not MYVALID Then Exit Sub
If MyReplace Then
    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
    CardTable.Seek "=", xDoc_No.Text
    MyLoad
Else
    CmdUndo_Click
    MsgBox "·„   „ ⁄„·Ì… «· ⁄œÌ· »‰Ã«Õ"
End If
End Sub
Private Sub CmdUndo_Click()
If CardTable.RecordCount = 0 Then
    mydefine
    Exit Sub
End If

CardTable.Seek "=", xDoc_No.Text
If CardTable.NoMatch Then
    CardTable.MoveLast
End If
MyLoad
End Sub
Private Sub cmdLoadMem_click()
rsMember.Seek "=", xCode.Text
If Not rsMember.NoMatch Then
    Load Member
'    Member.LoadOut xCode.Text
    Member.Show 1
End If
End Sub

Private Sub Command3_Click()
 ' ÷»ÿ «·„ÿ«·»«  «·”«»ﬁ…
'Set myItemTable = mydb.OpenRecordset("file2_10", dbOpenSnapshot)
'
'myItemTable.FindFirst "islateitem"
'CLATE = myItemTable.code
'
'myItemTable.FindFirst "overdue"
'COVERDUE = myItemTable.code
'
'mydb.Execute "UPDATE file2_30 INNER JOIN file2_20 ON file2_30.DOC_NO = file2_20.doc_no SET file2_20.OVERDUE = [FILE2_30].[TOTAL] " & _
'             " WHERE file2_30.ITEM = " & COVERDUE
'
'mydb.Execute "UPDATE file2_30 INNER JOIN file2_20 ON file2_30.DOC_NO = file2_20.doc_no SET file2_20.LATE = [FILE2_30].[TOTAL] " & _
'             " WHERE file2_30.ITEM = " & CLATE
'
'mydb.Execute "DELETE file2_3.* FROM FILE2_30 WHERE ITEM = " & CLATE
'mydb.Execute "DELETE file2_3.* FROM FILE2_30 WHERE ITEM = " & COVERDUE
Set Sumtable = mydb.OpenRecordset("Select doc_No,Sum(File2_30.Total) as sumoftotal from File2_30 where Year = 0 " & _
                                  "Group by Doc_No")
Do
    mydb.Execute "update file2_20 Set File2_20.total = iif(isNull(File2_20.OverDue),0,File2_20.OverDue) + iif(isNull(File2_20.Late),0,File2_20.Late) + " & TurnValue(Sumtable!SUMOFTOTAL, Null, 0) & _
                  ", File2_20.total0 = " & TurnValue(Sumtable!SUMOFTOTAL, Null, 0) & _
                  " where file2_20.Doc_No = " & MyParn(Sumtable!doc_no)
    Sumtable.MoveNext
Loop Until Sumtable.EOF

mydb.Execute "update FILE2_20 SET file2_20.[type] = -1 where isNull(File2_20.[Type])"
mydb.Execute "update FILE2_20 SET file2_20.type = -1 where File2_20.Type = 9"
MsgBox " „ «·÷»ÿ »‰Ã«Õ"
End Sub
Private Sub Command1_Click()
bNoRelation = True
If xType.Text = "" Then
    MsgBox "‰Ê⁄ «·„ÿ«·»… €Ì— „”Ã·"
    Exit Sub
End If
nYears = yearsToAdd
CheckOver
If nYears >= 0 Then additems nYears
End Sub
Private Sub Form_Load()
Frame11.Visible = bShowRep
Set tPayItem = mydb.OpenRecordset("file2_30")
Set paidTypeTable = mydb.OpenRecordset("select * From File0_50 order by code")
Set CardTable = mydb.OpenRecordset("File2_20")
Set ItemTable = mydb.OpenRecordset("select * from FILE2_10 Order by CODE ", dbOpenDynaset)
Set itemSectionTable = mydb.OpenRecordset("SELECT FILE2_10.CODE, FILE2_10.DESCA, FILE2_40.VALUE,FILE2_40.DISCOUNT,FILE2_40.TYPE,FILE2_40.SECTION,FILE2_40.YEAR " & _
                      " from file2_10 inner join file2_40 ON FILE2_10.CODE = FILE2_40.item ", dbOpenSnapshot)

Set SectionTable = mydb.OpenRecordset("select * from FILE0_10 ", dbOpenDynaset)
Set agetable = mydb.OpenRecordset("File0_30", dbOpenSnapshot)
Set MeetingTable = mydb.OpenRecordset("FILE1_12")

Data2.DatabaseName = MdbPath
Data2.RecordSource = "Select Code,DescA from file0_00 Where Flag = " & MEM_COMPANY
xCompany.BoundColumn = "Code"
xCompany.ListField = "DESCA"


Data1.DatabaseName = MdbPath
Data1.RecordSource = "Select Code,DescA from file0_50"
xType.BoundColumn = "Code"
xType.ListField = "DESCA"

MeetingTable.index = "ndx1"

CardTable.index = "ndx2"
tPayItem.index = "ndxDoc_No"

For I = 0 To 3
    Grid1(I).Cols = 8
    Grid1(I).rows = 2
    Grid1(I).TextMatrix(0, 0) = "ﬂÊœ"
    Grid1(I).TextMatrix(0, 1) = "«·»Ì«‰"
    Grid1(I).TextMatrix(0, 2) = "«·ﬁÌ„…"
    Grid1(I).TextMatrix(0, 3) = "⁄œœ"
    Grid1(I).TextMatrix(0, 4) = "Œ’„"
    Grid1(I).TextMatrix(0, 5) = "€—«„…  √ŒÌ—"
    Grid1(I).TextMatrix(0, 6) = "≈Ã„«·Ï"
    Grid1(I).ColHidden(7) = True
    
    Grid1(I).ColFormat(2) = "##0.00"
    Grid1(I).ColFormat(6) = "##0.00"
    
    For i2 = 0 To 7
        Grid1(I).ColAlignment(i2) = flexAlignRightCenter
    Next i2
    sSystem = xDate.Tag
    Grid1(I).ColWidth(0) = 500
    Grid1(I).ColWidth(1) = 3000
    Grid1(I).ColWidth(2) = 1000
    Grid1(I).ColWidth(3) = 1000
    Grid1(I).ColWidth(4) = 1000
    Grid1(I).ColWidth(5) = 1000
    Grid1(I).ColWidth(6) = 1000
    Grid1(I).FixedRows = 1
    Grid1(I).Editable = flexEDKbdMouse
Next
sSystem = xDate.Tag
If CardTable.RecordCount > 0 Then
    CardTable.MoveLast
    MyLoad
Else
    mydefine
    xDoc_No.Text = "000001"
End If
End Sub

Private Sub Grid1_StartEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Grid1(index).Row = Grid1(index).rows - 1 Then
    Grid1(index).AddItem ""
End If
End Sub

Private Sub xCode_LostFocus()
LoadMember
If xStop.Value = 1 Then
    MsgBox "«·⁄„Ì· „ Êﬁ›"
    If xCode.Enabled Or bAdd Then
        xCode.Text = ""
        xDescA.Caption = ""
        xSectionName.Caption = ""
        xLastYear.Caption = ""
        xDebit_String.Caption = ""
        xStop.Value = 0
    End If
End If
End Sub

Private Sub xCompany_Change()
cmdComp.Enabled = xCompany.MatchedWithList
cmdComp2.Enabled = xCompany.MatchedWithList
cmdcomp3.Enabled = xCompany.MatchedWithList
End Sub

Private Sub xDate_LostFocus()
LoadTabCaption
If IsDate(xDate.Text) Then
    If (xType.BoundText = 1 And MyEmpty(xYear.Text)) Or xType.BoundText <> 1 Then
        xYear.Text = PaidYear(xDate.Text)
    End If
Else
        xYear.Text = ""
End If
End Sub

Private Sub xDoc_No_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID() As Boolean
'If Format(Date, "yyyy-mm-dd") > sSystem Then Exit Function
If xCode = "" Then Exit Function
rsMember.Seek "=", xCode.Text
If rsMember.NoMatch Then Exit Function
If Not IsDate(xDate.Text) Then Exit Function
If xType.BoundText = 1 Then
'    If IsNull(RetLast(Val(xCode.Text))) Then
'        MsgBox "·« Ì„ﬂ‰ Õ›Ÿ ⁄÷ÊÌ… «·⁄÷Ê «–« «‰Â ·„ Ì”œœ „‰ ﬁ»·"
'        Exit Function
'    End If
'    If PaidYear(xDate.Text) - 1 > PaidYear(RetLast(rsMember!code)) Then
'        MsgBox "⁄·Ì «·⁄„Ì·  ”œÌœ«  ”«»ﬁ… ·« Ì„ﬂ‰ Õ›Ÿ ⁄÷ÊÌ Â"
'        Exit Function
'    End If
    
'    If Val(xYear.Text) = 0 Then
'        MsgBox "·„ Ì „ «œ—«Ã ”‰… ”œ«œ ..„‰ ›÷·ﬂ «œ—Ã ”‰… «·”œ«œ"
'        Exit Function
'    End If
End If

For nTab = 0 To 3
    If Tab1.TabVisible(nTab) Then
        nLastrow = LastRow(Grid1(nTab))
        For nrow = 1 To nLastrow
            If Grid1(nTab).TextMatrix(nrow, 0) = "" Then
                MsgBox "«·»‰œ €Ì— „ÊÃÊœ"
                Exit Function
            End If
        Next
    End If
Next

If Grid1(0).rows < 3 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·Õ›ŸÂ«"
    Exit Function
End If
'If IsNumeric(xNo.Text) Then
'    Dim cDoc_no As String
'    cDoc_no = GetDesca("select doc_no from file2_20 where [no] = " & xNo.Text & " and year = " & Val(xYear.Text) & " and doc_no <> " & MyParn(xDoc_No.Text))
'    If cDoc_no <> "" Then
'        MsgBox "—ﬁ„ «·ﬁ”Ì„… „ÊÃÊœ „‰ ﬁ»· ›Ï „” ‰œ —ﬁ„ " & cDoc_no
'    End If
'End If
MYVALID = True
End Function
Private Sub MyLoad()
On Error Resume Next
xCode.Text = TurnValue(CardTable.CODE, Null, "")
LoadMember
xDoc_No.Text = TurnValue(CardTable.doc_no, Null, "")
xDate.Text = Format(CardTable!Date, "DD-MM-YYYY")
xNo.Text = TurnValue(CardTable!NO, Null, "")
xType.BoundText = TurnValue(CardTable!Type)
xOverDue.Text = Format(CardTable!OverDue, "Fixed")
xLate.Text = Format(CardTable!LATE, "Fixed")
xYear.Text = TurnValue(CardTable!Year, Null, "")
xDebit_String.Caption = ""
'If Err.Number <> 0 Then
'    MsgBox " „ Õ–› «·”Ã· „‰ ﬁ»· „” Œœ„ √Œ— !!”Ì „ ⁄„·  Õ„Ì· ·«Œ—  ⁄œÌ·"
'    If CardTable.RecordCount > 0 Then
'        CardTable.FindFirst "Code = " & MyParn(xCode.Text)
'        If CardTable.NoMatch Then CardTable.MoveFirst
'        myload
'    Else
'        myDefine
'    End If
'End If
LoadTabCaption
Handlecontrols LoadMode
fillgrd
Calctotals
End Sub
Private Sub mydefine()
xTotal0.Caption = ""
xTotal.Caption = ""
xCode.Text = ""
xDate.Text = Format(Date, "dd-mm-yyyy")
xYear.Text = PaidYear(Date)
xDescA.Caption = ""
xNo.Text = ""
xLate.Text = ""
xOverDue.Text = ""
xLastYear.Caption = ""
xSectionName.Caption = ""
xType.BoundText = 0
xDebit_String.Caption = ""
For nTab = 0 To 3
    Grid1(nTab).rows = 1
    Grid1(nTab).AddItem ""
    If nTab > 0 Then
        LoadTabCaption
        Tab1.TabVisible(nTab) = False
    End If
Next
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdNext.Enabled = (nMode = LoadMode)
CmdPrevious.Enabled = (nMode = LoadMode)
CmdFirst.Enabled = (nMode = LoadMode)
CmdLast.Enabled = (nMode = LoadMode)
xYear.Enabled = (xType.BoundText = 1)
xCode.Enabled = Not (nMode = LoadMode)
xDoc_No.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub xDoc_No_LostFocus()
CardTable.Seek "=", xDoc_No.Text
If Not CardTable.NoMatch Then MyLoad
End Sub
Private Sub CmdDetial_Click()
bNoRelation = False
If xType.Text = "" Then
    MsgBox "‰Ê⁄ «·„ÿ«·»… €Ì— „”Ã·"
    Exit Sub
End If
'nYears = yearsToAdd
'If nYears >= 0 Then
'    MYDEFINETAB
'    CheckOver
'    Additems nYears
'    If xType.BoundText = 1 Then addMulti
'End If

Dim aYears As Variant
aYears = aPaidYears(Format(xDate.Text, "yyyy-mm-dd"), RetLast(xCode.Text, 2) & "")
MYDEFINETAB
If xType.BoundText = "0" Or xType.BoundText = "1" Then
    If IsEmpty(aYears) Then
        MsgBox "·Ì” ⁄·Ì «·⁄÷Ê ”‰Ê«  ”«»ﬁ…"
        For I = 1 To 3
            Tab1.TabVisible(I) = False
        Next
    ElseIf UBound(aYears) >= 3 Then
        MsgBox "«·⁄÷Ê  ⁄œÌ ”‰Ê«  «·”œ«œ"
    Else
        CheckOver
        additems (aYears)
        If retFlag(aYears(0), "rate") = "1" Then
            xDebit_String.Caption = "⁄·ÌÂ ”‰…"
        Else
            xDebit_String.Caption = "⁄·ÌÂ ”‰… Ê‰’"
        End If
        If xType.BoundText = 1 Then addMulti
    End If
ElseIf xType.BoundText = "2" Then
    If IsEmpty(aYears) Then
        MsgBox "«·⁄÷Ê „”œœ „‰ ﬁ»· Ê·Ì” ⁄·ÌÂ ”œ«œ"
        For I = 1 To 3
            Tab1.TabVisible(I) = False
        Next
    ElseIf UBound(aYears) <> 0 Then
        MsgBox "«·⁄÷Ê  ⁄·ÌÂ «ﬂÀ— „‰ ”‰…"
    Else
        If Not retFlag(aYears(0), "is_New", False) Then
            MsgBox "«·⁄÷Ê „”œœ „‰ ﬁ»·"
            Exit Sub
        End If
        CheckOver
        additems (aYears)
        If retFlag(aYears(0), "rate") = "1" Then xDebit_String.Caption = "⁄·ÌÂ ”‰…" Else xDebit_String.Caption = "⁄·ÌÂ ”‰… Ê‰’"
    End If
ElseIf xType.BoundText = "3" Then
    If Not IsEmpty(aYears) Then
        MsgBox "«·⁄÷Ê ·„ Ì”œœ «Œ— ”œ«œ"
        Exit Sub
    End If
    Dim aPaid As Variant
    aPaid = AddFlag(Empty, "tab", 0)
    aPaid = AddFlag(aPaid, "year", PaidYear(xDate.Text))
    aPaid = AddFlag(aPaid, "year_string", Year(xDate.Text))
    aPaid = AddFlag(aPaid, "rate", 1)
    aYears = AddFlag(aYears, aPaid)
    CheckOver
    additems (aYears)
    If retFlag(aYears(0), "rate") = "1" Then xDebit_String.Caption = "⁄·ÌÂ ”‰…"
End If
End Sub
Private Sub Fixmembers()
Dim I As Integer, nRecord As Integer
Prog1.Visible = True
nRecordcount = rsMember.RecordCount
rsMember.MoveFirst
Do
    I = I + 1
    'FixMemberPaid (rsMember!code)
    rsMember.Edit
    rsMember!Year = RetLast(rsMember!CODE, 0)
    rsMember!IsSave = RetLast(rsMember!CODE, 1)
    rsMember!datelast = RetLast(rsMember!CODE, 2)
    rsMember!doc_no = RetLast(rsMember!CODE, 3)
    rsMember.Update
    rsMember.MoveNext
    If Prog1.Value <> Int(I / nRecordcount * 100) Then Prog1.Value = Int(I / nRecordcount * 100)
Loop Until rsMember.EOF
Prog1.Visible = False
End Sub
Private Sub CmdPrint_Click()
doprint
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
mydefine
xDoc_No.Text = "000001"
End Sub
Private Sub editProc()
formMode = EditMode
MyLoad
Handlecontrols (EditMode)
End Sub
Sub addProc()
formMode = addMode
Handlecontrols addMode
mydefine
xDoc_No.Text = IncRec(myLastField(CardTable, "DOC_NO"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'memItemTable.Close
'Set memItemTable = Nothing
Unload Me
Set paid = Nothing
Unload Search2
If isopen Then Unload Search
End Sub
Sub fillgrd()
Dim nTotal As Double

For I = 0 To 3
    If I > 0 Then Tab1.TabVisible(I) = False
    Grid1(I).rows = 1
Next

With tPayItem
If .RecordCount = 0 Then Exit Sub
.Seek ">=", xDoc_No.Text, 0
If .NoMatch Then
    Grid1(0).AddItem ""
    Exit Sub
End If
If .doc_no <> xDoc_No.Text Then Exit Sub

I = 1
nYear = -1
nTab = -1
Do While !doc_no = xDoc_No.Text
    If !Year <> nYear Then
        nTab = nTab + 1
        nYear = .Year
        Tab1.TabVisible(nTab) = True
        Grid1(nTab).rows = 1
    End If
    Grid1(nTab).AddItem ""
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = TurnValue(.Item, Null, "")
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = RetDescA(ItemTable, .Item)
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = aTurnValue(!Value, Array(Null, 0), "")
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = aTurnValue(!Number, Array(Null, 0), "")
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = aTurnValue(!DISCOUNT, Array(Null, 0), "")
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = aTurnValue(!Rate, Array(Null, 0), "")
    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 7) = aTurnValue(!Note, Array(Null, 0), "")
    .MoveNext
    If .EOF Then Exit Do
Loop
For I = 0 To 3
    Grid1(I).AddItem ""
Next
End With
End Sub
Private Sub Grid1_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
With Grid1(index)
If Row = 0 Then Exit Sub
If Col = 0 Or Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Then
    Calctotals
End If
End With
End Sub
Private Sub Grid1_EnterCell(index As Integer)
If Grid1(index).Col = 6 Then
    Grid1(index).Editable = flexEDKbd
Else
    Grid1(index).Editable = flexEDKbdMouse
End If
End Sub
Private Sub grid1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
With Grid1(index)
If Grid1(index).Col = 0 And KeyCode = 112 Then
    CODELOOKUP
End If
If KeyCode = 46 Then
    If Grid1(index).Row = Grid1(index).rows - 1 Then Exit Sub
    If MsgBox("Õ–› »‰œ _ Â· «‰  „ √ﬂœ ø", 1 + 256) = vbOK Then
        Grid1(index).RemoveItem Grid1(index).Row
        Calctotals
    End If
End If
End With
End Sub
Private Sub grid1_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid1(index)
If .EditText = "" Then Exit Sub
If Col = 0 Then
    ItemTable.FindFirst "CODE = " & .EditText
    If Not ItemTable.NoMatch Then
         .TextMatrix(.Row, 1) = ItemTable!DESCA
         .TextMatrix(.Row, 2) = RetValue(.EditText, index)
         .TextMatrix(.Row, 3) = 1
         .TextMatrix(.Row, 4) = RetValue(.EditText, index, 1)
    Else
         MsgBox "«·»‰œ €Ì— „ÊÃÊœ ·Â–Â «·›∆… «Ê ·Ì” „ÊÃÊœ ⁄·Ì «·«ÿ·«ﬁ"
         Cancel = True
    End If
End If
End With
End Sub
Private Sub xCode_KeyPress(KeyAscii As Integer)
KeyAscii = RetNumber(KeyAscii, False)
End Sub
Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(2, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select code,[desca],IIF(STOP,'„ Êﬁ›','') from file1_10"
    Generalarray(2) = "Order by FILE1_10.CODE"
    Generalarray(3) = 7000
    Generalarray(5) = True
    
    listarray(0, 0) = "«·«”„"
    listarray(0, 1) = "(%%DESCA%%)"
    
    
    GrdArray(0, 0) = "—ﬁ„ «·⁄÷ÊÌ…"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·«”„"
    GrdArray(1, 1) = 6000
    
    GrdArray(2, 0) = "«·Õ«·…"
    GrdArray(2, 1) = 1000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    oSearchMember.bEnter = True
    oSearchMember.Caption = "«” ⁄·«„ «·«⁄÷«¡"
    oSearchMember.Show 1
End If
End Sub
Private Sub grid1_KeyUpEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If Grid1(index).Col = 0 Then
    If KeyCode = 112 Then CODELOOKUP
End If
End Sub
Private Function RetLate(pTab, nYears) As Double
If xType.BoundText <> 0 Then Exit Function
If IsNull(RetLast(xCode.Text)) Then Exit Function
dDatePaid = paidDate(PaidYear(xDate.Text))
nDays = DateDiff("d", dDatePaid, xDate.Text)
Select Case pTab
Case 0
    'If nLateDays > 0 Then RetLate = IIf(nDays >= nLateDays, 50, 0)
Case 1
    RetLate = 50
Case 2
    RetLate = 100
Case 3
    RetLate = 200
End Select
End Function
Private Sub Calctotals()
Dim nTotal1 As Double, nTotal As Double, nOverDue, nLate

' «·«‘ —«ﬂ «·”‰ÊÌ
For I = 1 To Grid1(0).rows - 1
' Õ”«» «·”‰… «·Õ«·Ì… »œÊ‰ €—«„…
    nTotalRow = Val(Grid1(0).TextMatrix(I, 2)) * _
                Val(Grid1(0).TextMatrix(I, 3)) * _
                (1 - Val(Grid1(0).TextMatrix(I, 4)) / 100)
    Grid1(0).TextMatrix(I, 6) = TurnValue(nTotalRow, 0, "")
    
    nTotal1 = nTotalRow + nTotal1

' Õ”«» «·€—«„…
'    If nSystem <> ETAHAD_SYSTEM Then
'        nLate = nLate + ( _
'                   Val(Grid1(0).TextMatrix(I, 2)) * _
'                   Val(Grid1(0).TextMatrix(I, 3)) * _
'                  (1 - Val(Grid1(0).TextMatrix(I, 4)) / 100) * _
'                  (Val(Grid1(0).TextMatrix(I, 5)) / 100))
'    Else
'        nLate = nLate + ( _
'                   Val(Grid1(0).TextMatrix(I, 2)) * _
'                   Val(Grid1(0).TextMatrix(I, 3)) * _
'                  (Val(Grid1(0).TextMatrix(I, 5)) / 100))
'    End If
Next

For nTab = 1 To 3
    For I = 1 To Grid1(nTab).rows - 1
        nTotalRow = Val(Grid1(nTab).TextMatrix(I, 2)) * _
                                      Val(Grid1(nTab).TextMatrix(I, 3)) * _
                                      (1 - Val(Grid1(nTab).TextMatrix(I, 4)) / 100)
        Grid1(nTab).TextMatrix(I, 6) = TurnValue(nTotalRow, 0, "")
    Next
Next
xTotal0.Caption = Format(nTotal1, "Fixed")

nLate = CalcLate
nOverDue = CalcOverDue
If nOverDue <> 0 Then
    xOverDue.Text = Format(nOverDue, "Fixed")
End If

'If nLate <> 0 Then
    xLate.Text = Format(nLate, "Fixed")
'End If

xTotal.Caption = Format(Val(xTotal0.Caption) + Val(xOverDue.Text) + Val(xLate.Text), "Fixed")
End Sub
Private Function RetPer(cString) As String
RetPer = IIf(cString = 0, "", cString & "%")
End Function
Sub FixMemberPaid(nMember)
With rsMember
    rsMember.Seek "=", nMember
    .Edit
    !Year = RetLast(nMember, 0)
    !IsSave = RetLast(nMember, 1)
    !datelast = RetLast(nMember, 2)
    !doc_no = RetLast(nMember, 3)
    rsMember.Update
End With
End Sub
Private Sub Command2_Click() ' Fixing Functions
nDays = InputBox("«œŒ· «Ì«„ «·€—«„…", "√Ì«„ «·€—«„…")
If Val(nDays) > 0 Then
    mydb.Execute "update setting set days = " & nDays
    nLateDays = nDays
End If
End Sub
Private Sub FixDoc()
Dim nYear As Integer, DocTable As Recordset, nAdd As Integer
Set DocTable = mydb.OpenRecordset("Select * From File2_30 Where Year > 10 Order by Doc_No,Year")
DocTable.MoveLast
nRecords = DocTable.RecordCount
DocTable.MoveFirst
nYear = DocTable!Year
cDoc = DocTable!doc_no
nAdd = 0
Prog1.Value = 0
Prog1.Visible = True
I = 0
Do
    I = I + 1
    DocTable.Edit
    If DocTable.doc_no <> cDoc Then
        cDoc = DocTable!doc_no
        nYear = DocTable.Year
        nAdd = 0
    ElseIf DocTable.doc_no = cDoc And nYear <> DocTable!Year Then
        nYear = DocTable.Year
        nAdd = nAdd + 1
    End If
    DocTable.Year = nAdd
    DocTable.Update
    If Prog1.Value <> Int(I / nRecords * 100) Then Prog1.Value = Int(I / nRecords * 100)
    DocTable.MoveNext
Loop Until DocTable.EOF
End Sub
Private Sub PrintCash()
tempdb.Execute "delete * from temp"
Set temptable = tempdb.OpenRecordset("select * from Temp")
Set sourcetable = mydb.OpenRecordset("Select * From file2_30 Where " & _
                                     "Doc_No = " & MyParn(xDoc_No.Text) & _
                                     " and Year = 0 " & _
                                     " Order by Item ")
If Grid1(0).rows = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    Exit Sub
End If

If Not IsDate(xDate.Text) Then
   MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
   Exit Sub
End If
With Grid1(0)
nLastrow = LastRow(Grid1(0))
For I = 1 To nLastrow
    temptable.AddNew
    temptable!str1 = "„œ›Ê⁄«  Œ·«· ⁄«„ " & PaidString(Year(xDate.Text))
    temptable!str2 = RetFind(SectionTable, "Code", "DescA", RetMember(xCode.Text, "Section"))
    temptable!Str3 = .TextMatrix(I, 1)
    temptable!str11 = TurnValue(xDescA.Caption, "", Null)
    temptable!str12 = Chr(254) & xCode.Text
    temptable.str13 = TurnValue(RetMember(xCode.Text, Format("DateLast", "dd-mm-yyyy")), "", Null)
    temptable.str14 = Chr(254) & TurnValue(RetMember(xCode.Text, "Address"), "", Null)
    temptable!val1 = .TextMatrix(I, 2)
    temptable!val2 = .TextMatrix(I, 3)
    temptable!Val3 = TurnValue(.TextMatrix(I, 4), "", Null)
    temptable!val4 = TurnValue(.TextMatrix(I, 6), "", Null)
    temptable.Update
Next
End With
myws.BeginTrans
myws.CommitTrans
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\rpt\rpt02_1.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1: tempdb.Execute "Delete * from temp"
End Sub
Private Function addLocker(nType) As Double
Dim countTable As Recordset
Set countTable = mydb.OpenRecordset("Select Count(Member) as CountOfMember From File6_10 Where " & _
                                     " Member = " & xCode.Text & " and Type = " & MyParn(nType) & _
                                     " Group By Member ")
If countTable.RecordCount > 0 Then
    addLocker = countTable!CountOfMember
End If
End Function
Private Function addyacht(nType) As Double
Dim countTable As Recordset
Set countTable = mydb.OpenRecordset("Select Count(Member) as CountOfMember From File6_20 Where " & _
                                     " Member = " & xCode.Text & " and Type = " & MyParn(nType) & _
                                     " Group By Member ")
If countTable.RecordCount > 0 Then
    addyacht = countTable!CountOfMember
End If
End Function
Private Sub additems(aYears As Variant)
Dim nMeetValueMem As Double, nMeetValueWife As Double, bAsNew As Boolean, bDebit As Boolean, nFirstYear As Long, nlateValue As Double
Dim nYearPaid As Long
If xType.Text <> "" Then
    paidTypeTable.FindFirst "CODE = " & xType.BoundText
    bAsNew = paidTypeTable.ASNEW
End If
bDebit = RetMember(xCode.Text, "DEBIT")
nFirstYear = retFlag(aYears(0), "year")
nYears = UBound(aYears)
'If nYears = 1 Then nYears = 0
For nTab = 0 To nYears
    'nyearPaid = PaidYear(xDate.Text) - nTab
    nYearPaid = retFlag(aYears(nTab), "year")
    cString = "SELECT FILE2_10.CODE, FILE2_10.DESCA, FILE2_10.ALLMEMBER, FILE2_10.LATE, FILE2_10.RELATION," & _
          " FILE2_10.ISMEMBER, FILE2_10.AGE1, FILE2_10.AGE2, FILE2_10.SEX, FILE2_10.Locker, FILE2_10.yacht, " & _
          " FILE2_10.BASICDIED,FILE2_10.BASICOLD, FILE2_10.BASICNEW, FILE2_10.MEETING, FILE2_10.DAYS, FILE2_10.NORATE, " & _
          " FILE2_40.value, FILE2_40.Discount " & _
          " FROM FILE2_10 INNER JOIN FILE2_40 ON FILE2_10.CODE = FILE2_40.item " & _
          " WHERE FILE2_40.TYPE = " & xType.BoundText & _
          " AND FILE2_40.BASIC " & _
          " AND FILE2_40.YEAR = " & nYearPaid & _
          " AND SECTION =  " & MyParn(RetMember(xCode.Text, "Section"))
    
    If RetMember(xCode.Text, "Died") Then
        cString = cString & " and BASICDIED"
    End If
    
    cString = cString & " ORDER BY FILE2_10.CODE"
    
    Set MemItemTable = mydb.OpenRecordset(cString)
    
    With MemItemTable
    If MemItemTable.RecordCount = 0 Then
        MsgBox "·Ì” ··›∆… «Ì… »‰Êœ ·«÷«› Â« ›Ï ”‰… " & nYearPaid
        GoTo LastNext
    End If

    Tab1.TabVisible(nTab) = True
    Grid1(nTab).rows = 1
    Do
        If MemItemTable!isMember Then
          nNewMeeting = 0
          If AddMember(nTab) Then
               Grid1(nTab).AddItem ""
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = 1
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = TurnValue(IIf(MemItemTable!LATE, RetLate(nTab, nYears), ""), 0, "")
               
               nlateValue = mRound((Val(MemItemTable!Value & "") * (61.56 / 184.68)) * (1 - (Val(MemItemTable!DISCOUNT & "") / 100)), 2)
               Calctotals
               If nTab = 0 Then
                    If nNewMeeting <> 0 Then
                        nMeetValueMem = nMeetValueMem + (nNewMeeting * Val(Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 6)))
                    End If
               End If
           End If
           GoTo LastLoop
        End If
        
        If (Not IsNull(!Relation)) And !meeting = False Then
          If bNoRelation Then GoTo LastLoop
          nNewMeeting = 0
          nNumber = addRelation(!Relation, nTab)
          If nNumber > 0 Then
               Grid1(nTab).AddItem ""
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = nNumber
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = TurnValue(IIf(MemItemTable!LATE, RetLate(nTab, nYears), 0), 0, "")
               Calctotals
               If nTab = 0 And !Relation & "" = 1 Then
                    If nNewMeeting <> 0 Then
                        nMeetValueWife = nMeetValueWife + (nNewMeeting * Val(Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 6)))
                    End If
               End If
           End If
           GoTo LastLoop
        End If
           
        If Not IsNull(!Locker) Then
           nNumber = addLocker(MemItemTable!Locker)
           If nNumber <> 0 Then
               Grid1(nTab).AddItem ""
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = nNumber
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
               Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
           End If
           GoTo LastLoop
        End If
                                          
        If !meeting Then
'            aret = RetLastNew(xCode.Text)
'            If Format(retFlag(aret, "date"), "yyyy-mm-dd") <= "2015-02-10" Then
                If nTab = 0 Then
                    If IsNull(!Relation) And nMeetValueMem <> 0 Then
                        Grid1(nTab).AddItem ""
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = nMeetValueMem
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = 1
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
                    ElseIf (Not IsNull(!Relation)) And nMeetValueWife <> 0 Then
                        Grid1(nTab).AddItem ""
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = nMeetValueWife
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = 1 * retFlag(aYears(nTab), "rate")
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
                        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
                    End If
                End If
'            End If
            GoTo LastLoop
        End If
        
        If !BasicNew Then
            If IsNull(RetLast(xCode.Text)) Or bAsNew Then
                Grid1(nTab).AddItem ""
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = IIf(!AllMember, AllCount(nTab), 1)
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), 0)
                GoTo LastLoop
            End If
            If Not !BASICOLD Then GoTo LastLoop
        End If
        
        If !BASICOLD Then
            If Not IsNull(RetLast(xCode.Text)) Then
                Grid1(nTab).AddItem ""
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = IIf(!AllMember, AllCount(nTab), 1)
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
            End If
            GoTo LastLoop
        End If
        
        If !BasicDied Then
            If RetMember(xCode.Text, "Died") Then
                Grid1(nTab).AddItem ""
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = IIf(!AllMember, AllCount(nTab), 1)
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
                Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
            End If
            GoTo LastLoop
        End If

        If (Not IsNull(!days)) Then
            If DateDiff("d", paidDate(retFlag(aYears(nTab), "year")), xDate.Text) >= !days And nTab = 0 And Not IsNull(RetLast(xCode.Text)) Then
                If AddLate(nTab) Then
                    Grid1(nTab).AddItem ""
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = nlateValue * retFlag(aYears(nTab), "rate")
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = 1
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = ""
                    Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = ""
                End If
            End If
            GoTo LastLoop
        End If

        Grid1(nTab).AddItem ""
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 0) = MemItemTable!CODE
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 1) = MemItemTable!DESCA & ""
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 2) = Val(MemItemTable!Value & "") * retFlag(aYears(nTab), "rate")
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 3) = IIf(!AllMember, AllCount(nTab), 1)
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 4) = TurnValue(MemItemTable!DISCOUNT, Null, "")
        Grid1(nTab).TextMatrix(Grid1(nTab).rows - 1, 5) = IIf(MemItemTable!LATE, RetLate(nTab, nYears), "")
LastLoop:
        MemItemTable.MoveNext
    Loop Until MemItemTable.EOF
    Grid1(nTab).AddItem ""
LastNext:
End With
Next

Calctotals
Exit Sub
myerror:
MsgBox "Â‰«ﬂ Œÿ« „« «À‰«¡ ÷»ÿ «·»Ì«‰«  „‰ ›÷·ﬂ —«Ã⁄ »Ì«‰« ﬂ"
End Sub

Private Sub xLate_LostFocus()
'If Val(xLate) = 0 And CalcLate <> 0 Then
'    MsgBox "”Ì „ Õ–› €—«„«  «· «ŒÌ—"
'    MyDefineLate
'End If
'CalcTotals
Calctotals
End Sub
Private Sub xOverDue_LostFocus()
'If Val(xOverDue) = 0 Then
'     MsgBox "”Ì „ Õ–› «·”‰Ê«  «·„ «Œ—…"
'    MyDefineOverDue
'End If
Calctotals
End Sub
Private Function AddRelSep() As Boolean
If IsNull(MemItemTable.Relation) Then
    Exit Function
End If
AddRelSep = True
'agetable.FindFirst "code = " & memItemTable!code
'AddRelSep = (agetable.Relation = 1)
End Function
Private Sub LoadMember()
xDescA.Caption = ""
xSectionName.Caption = ""
xLastYear.Caption = ""
xDebit_String.Caption = ""

If xCode.Text = "" Then Exit Sub

rsMember.Seek "=", xCode.Text
If rsMember.NoMatch Then
    xDescA = ""
    xLastYear = ""
    xSectionName.Caption = ""
    xStop.Value = 0
    Exit Sub
End If
xDescA.Caption = rsMember!DESCA
xLastYear.Caption = PaidString(xCode.Text)
xStop.Value = IIf(rsMember!Stop, 1, 0)


'xDebit_String.Caption = IIf(rsMember!Debit, "⁄·ÌÂ ”‰…", "⁄·ÌÂ ”‰… Ê‰’›")

If Not IsNull(RetMember(xCode.Text, "Section")) Then
    xSectionName.Caption = RetFind(SectionTable, "Code", "DescA", RetMember(xCode.Text, "Section"))
Else
    xSectionName.Caption = ""
End If
End Sub
Private Sub MyDefineLate()
For nTab = 0 To 3
    For I = 1 To Grid1(nTab).rows - 2
        Grid1(nTab).TextMatrix(I, 5) = ""
    Next
Next
End Sub
Private Sub MyDefineOverDue()
For nTab = 1 To 3
    Grid1(nTab).rows = 1
    Grid1(nTab).AddItem ""
    Tab1.TabVisible(nTab) = False
Next
End Sub
Private Function yearsToAdd() As Integer
yearsToAdd = -1
rsMember.Seek "=", xCode.Text
If rsMember.NoMatch Then
    MsgBox "ﬂÊœ «·⁄÷Ê €Ì— ’ÕÌÕ"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xCode.Text = "" Then
    MsgBox "ﬂÊœ «·⁄÷Ê €Ì— „”Ã·"
    Exit Function
End If

If IsNull(RetMember(xCode.Text, "Section")) Then
    MsgBox "›∆… «·⁄÷Ê €Ì— „”Ã·… ›Ï »Ì«‰«  «·⁄÷Ê"
    Exit Function
End If


If IsNull(RetLast(xCode.Text)) Then
    If xType.BoundText <> 1 Then
        yearsToAdd = 1
        Exit Function
    Else
        MsgBox "«·⁄÷Ê ÃœÌœ ·« Ì”„Õ »⁄„· „ÿ«·»… " & xType.Text
        Exit Function
    End If
End If

nYears = PaidYear(xDate.Text) - RetLast(xCode.Text)
If xType.BoundText = 0 Then
    If nYears <= 0 Then
        MsgBox "«·⁄÷Ê „”œœ Õ Ì ¬Œ— „Ê”„"
        Exit Function
    ElseIf nYears > 0 And nYears <= 4 Then
        yearsToAdd = nYears
        Exit Function
    Else
        MsgBox "⁄·Ì «·⁄÷Ê ”‰Ê«  «ﬂ»— „‰ ”‰Ê«  «··«∆Õ…"
        yearsToAdd = 1
    End If
End If

If xType.BoundText = 1 Then
    If nYears <= 1 Then
        yearsToAdd = 1
        nMulti = Val(xYear.Text) - RetLast(xCode.Text)
        nMutli = IIf(nMulti > 0, nMulti, 1)
        Exit Function
    Else
       MsgBox "⁄·Ì «·⁄÷Ê ”‰Ê«  ”«»ﬁ… ›·« Ì„ﬂ‰ Õ›Ÿ ⁄÷ÊÌ Â"
       Exit Function
    End If
End If
    
If xType.BoundText > 1 Then
    If nYears <= 0 Then
        yearsToAdd = 1
        Exit Function
    Else
       MsgBox "«·⁄÷Ê ·„ Ì”œœ «·”‰… «·Õ«·Ì… «Ê ”‰Ê«  ”«»ﬁ…"
       Exit Function
    End If
End If
End Function
Function CalcLate() As Double
For nTab = 0 To 3
    For I = 1 To Grid1(nTab).rows - 2
        If nSystem <> ETAHAD_SYSTEM Then
            nValue = (Val(Grid1(nTab).TextMatrix(I, 2)) * _
                       Val(Grid1(nTab).TextMatrix(I, 3)) * _
                      (1 - Val(Grid1(nTab).TextMatrix(I, 4)) / 100) * _
                      (Val(Grid1(nTab).TextMatrix(I, 5)) / 100))
            nValue = Format(nValue, "fixed")
            CalcLate = CalcLate + nValue
        Else
            nValue = ( _
                       Val(Grid1(nTab).TextMatrix(I, 2)) * _
                       Val(Grid1(nTab).TextMatrix(I, 3)) * _
                      (Val(Grid1(nTab).TextMatrix(I, 5)) / 100))
            If nValue = 29.7 Then nValue = 30
            CalcLate = CalcLate + nValue
        End If
    Next
Next
End Function
Function CalcLateTab(nTab) As Double
For I = 1 To Grid1(nTab).rows - 2
    nValue = (Val(Grid1(nTab).TextMatrix(I, 2)) * _
               Val(Grid1(nTab).TextMatrix(I, 3)) * _
              (1 - Val(Grid1(nTab).TextMatrix(I, 4)) / 100) * _
              (Val(Grid1(nTab).TextMatrix(I, 5)) / 100))
    nValue = Format(nValue, "fixed")
    CalcLateTab = CalcLateTab + nValue
Next
End Function

Private Function CalcOverDue() As Double
For nTab = 1 To 3
    For I = 1 To Grid1(nTab).rows - 2
        CalcOverDue = CalcOverDue + Val(Grid1(nTab).TextMatrix(I, 2)) * _
                                      Val(Grid1(nTab).TextMatrix(I, 3)) * _
                                      (1 - Val(Grid1(nTab).TextMatrix(I, 4)) / 100)
    Next
Next
End Function
Private Sub Calcrow(nrow, nIndex)
With Grid1(nIndex)
If .TextMatrix(nrow, 0) = "" Then
    .TextMatrix(nrow, 1) = ""
    Exit Sub
End If
End With
End Sub
Private Sub FixPaid()
Set PaidTable = mydb.OpenRecordset("Select * From File2_20", dbOpenDynaset)
PaidTable.MoveLast
Prog1.Visible = True
Prog1.Value = 0
nRecordcount = PaidTable.RecordCount
PaidTable.MoveFirst
Do
    PaidTable.Edit
    PaidTable!Year = PaidYear(PaidTable!Date)
    PaidTable.Update
    I = I + 1
    If Prog1.Value <> Int(I / nRecordcount * 100) Then Prog1.Value = IIf(Int(I / nRecordcount * 100) > 100, 100, Int(I / nRecordcount * 100))
    PaidTable.MoveNext
Loop Until PaidTable.EOF
Prog1.Visible = False
End Sub
Private Sub LoadTabCaption()
Dim nYear As Long
nYear = PaidYear(xDate.Text)
For nTab = 0 To 3
    If nYear = 2015 Then nYear = 2014
    Tab1.TabCaption(nTab) = paidYearString(nYear, True)
    nYear = nYear - 1
Next
End Sub
Private Function AddMember(nTab) As Boolean
If nTab = 0 Then
    dDate = IIf(nSystem = HORSE_SYSTEM, paidDate(PaidYear(xDate.Text)), xDate.Text)
Else
    dDate = paidDate(PaidYear(xDate.Text) - nTab)
End If

If RetMember(xCode.Text, "died") Then Exit Function
If IsNull(MemItemTable!age1) And IsNull(MemItemTable!age2) Then
    AddMember = True
    Exit Function
End If

If Not IsNull(MemItemTable!age1) Then
    If IsNull(RetMember(xCode.Text, "dateBirth")) Then Exit Function
    If Age(RetMember(xCode.Text, "dateBirth"), dDate) < MemItemTable!age1 Then Exit Function
End If

If Not IsNull(MemItemTable!age2) Then
    If IsNull(RetMember(xCode.Text, "dateBirth")) Then Exit Function
    If Age(RetMember(xCode.Text, "dateBirth"), dDate) > MemItemTable!age2 Then Exit Function
End If

If nTab = 0 Then
    nNewMeeting = Round((addMeet / 2), 2)
End If
AddMember = True
End Function
Private Function AddLate(nTab) As Boolean
If nTab <> 0 Then Exit Function
dDate = paidDate(PaidYear(xDate.Text) - nTab)
If IsNull(MemItemTable!age1) And IsNull(MemItemTable!age2) Then
    AddLate = True
    Exit Function
End If

If Not IsNull(MemItemTable!age1) Then
    If IsNull(RetMember(xCode.Text, "dateBirth")) Then Exit Function
    If Age(RetMember(xCode.Text, "dateBirth"), dDate) < MemItemTable!age1 Then Exit Function
End If

If Not IsNull(MemItemTable!age2) Then
    If IsNull(RetMember(xCode.Text, "dateBirth")) Then Exit Function
    If Age(RetMember(xCode.Text, "dateBirth"), dDate) > MemItemTable!age2 Then Exit Function
End If
AddLate = True
End Function
Private Function addRelation(cRelation, nTab) As Integer
If nTab = 0 Then
    dDate = IIf(nSystem = HORSE_SYSTEM, paidDate(PaidYear(xDate.Text)), xDate.Text)
Else
    dDate = paidDate(PaidYear(xDate.Text) - nTab)
End If

cString = "select * From File1_11 where Member = " & xCode.Text & " and relation = " & MyParn(cRelation)
If xType.BoundText = 3 Then
    nCode = GetDesca("Select file1_11.code from file1_11 where member = " & Val(xCode.Text) & " and Relation = '1' order by dateBegin Desc")
End If
If nCode <> "" Then cString = "select * From File1_11 where Member = " & xCode.Text & " and relation = " & MyParn(cRelation) & " and code = " & nCode
Set reltable = mydb.OpenRecordset(cString)
Do Until reltable.EOF
    If Not IsNull(MemItemTable!Sex) Then
        If IsNull(reltable.Sex) Then
            GoTo endLoop
        Else
            If (MemItemTable!Sex <> reltable!Sex) Then GoTo endLoop
        End If
    End If
    
    If IsNull(MemItemTable!age1) And IsNull(MemItemTable!age2) Then
        addRelation = addRelation + 1
        GoTo endLoop
    End If
    
    If Not IsNull(MemItemTable!age1) Then
        If IsNull(reltable!DateBirth) Then GoTo endLoop
        If Age(reltable!DateBirth, dDate) < MemItemTable!age1 Then GoTo endLoop
    End If
        
    If Not IsNull(MemItemTable!age2) Then
        If IsNull(reltable!DateBirth) Then GoTo endLoop
        If Age(reltable!DateBirth, dDate) > MemItemTable!age2 Then GoTo endLoop
    End If
       
    If MemItemTable!Relation = 1 And nTab = 0 Then
        nMeet = nMeet + addMeet(reltable!CODE & "")
        nRecord = nRecord + 1
    End If
       
    addRelation = addRelation + 1
endLoop:
    reltable.MoveNext
Loop
If nRecord > 0 Then
    nNewMeeting = Round((nMeet / nRecord) / 2, 2)
End If
End Function
Private Function addMeeting(nTab) As Integer
Dim wifeTable As Recordset
Dim MeetDateTable

' ·Ì” ··⁄÷Ê  «—ÌŒ ”œ«œ
If IsNull(RetLast(xCode.Text)) Then Exit Function

If nTab <> 0 Then Exit Function

nYear = RetLast(xCode.Text)

Set MeetDateTable = mydb.OpenRecordset("file1_13", dbOpenSnapshot)
If MeetDateTable.RecordCount = 0 Then Exit Function

' ·Õ”«»  «—ÌŒ «·Ã„⁄Ì… ⁄„Ê„Ì…
MeetDateTable.FindFirst " year = " & nYear
If MeetDateTable.NoMatch Then Exit Function


If Not IsNull(RetMember(xCode, "DateBegin")) Then
    ' ·„ Ì„— ⁄·Ì «·⁄÷Ê ”‰…
    If Age(RetMember(xCode.Text, "dateBegin"), MeetDateTable!Date) < 1 Then Exit Function
End If

If MeetingTable.RecordCount = 0 Then
    addMeeting = 1
Else
    MeetingTable.Seek "=", MeetDateTable!CODE, xCode.Text, Null
    addMeeting = IIf(MeetingTable.NoMatch, 1, 0)
End If

Set wifeTable = mydb.OpenRecordset("Select * From File1_11 Where Relation = '1' and Member = " & xCode.Text)
If wifeTable.RecordCount = 0 Then Exit Function
Do
   If Not IsNull(wifeTable!DateBirth) Then
        If Not IsNull(wifeTable!DateBirth) Then
            If Age(wifeTable!DateBirth, MeetDateTable!Date) < 21 Then GoTo myEndLoop
        End If
   End If
   
   If Not IsNull(wifeTable!datebegin) Then
        If Not IsNull(wifeTable!datebegin) Then
            If Age(wifeTable!datebegin, MeetDateTable!Date) < 1 Then GoTo myEndLoop
        End If
   End If

    If MeetingTable.RecordCount = 0 Then
        addMeeting = addMeeting + 1
    Else
        MeetingTable.Seek "=", MeetDateTable!CODE, xCode.Text, wifeTable!CODE
        addMeeting = IIf(MeetingTable.NoMatch, addMeeting + 1, addMeeting)
    End If
myEndLoop:
    wifeTable.MoveNext
Loop Until wifeTable.EOF
End Function
Private Function AllCount(nTab) As Integer
Dim countTable As Recordset
If xType.BoundText = 1 Or xType.BoundText = 3 Then
    AllCount = 1
    Exit Function
End If

'If xType.BoundText <> 0 Then
'    AllCount = CardCount(nTab)
'    Exit Function
'End If

If Not RetMember(xCode.Text, "died") Then AllCount = 1
Set countTable = mydb.OpenRecordset("Select * From File1_11 Where  Member = " & xCode.Text)
If countTable.RecordCount = 0 Then Exit Function
countTable.MoveLast
AllCount = AllCount + countTable.RecordCount
End Function
Private Function CardCount(nTab) As Integer
nCode = MemItemTable!CODE
MemItemTable.MoveFirst
With MemItemTable
Do
    If MemItemTable!isMember Then
        If AddMember(nTab) Then CardCount = CardCount + 1
        GoTo LastLoop
    End If

    If Not IsNull(!Relation) Then
        nNumber = addRelation(!Relation, nTab)
        CardCount = CardCount + nNumber
        GoTo LastLoop
    End If
LastLoop:
    MemItemTable.MoveNext
Loop Until MemItemTable.EOF
End With
MemItemTable.FindFirst "code = " & nCode
End Function
Private Sub doprint(Optional nDist As Long = 0)
tempdb.Execute "delete * from temp"
tempdb.Execute "delete * from temp2"
Set temptable = tempdb.OpenRecordset("Temp2", dbOpenDynaset)
If Grid1(0).rows = 0 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    Exit Sub
End If

If Not IsDate(xDate.Text) Then
   MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
   Exit Sub
End If
For nTab = 0 To 3
    With Grid1(nTab)
    cnames = RelNames
    For I = 1 To .rows - 2
        temptable.AddNew
        temptable!str17 = "„ÿ«·»… «·√⁄÷«¡ ⁄‰ ⁄«„ " & paidYearString(PaidYear(xDate.Text), True)
        temptable.str1 = xDescA.Caption
        ItemTable.FindFirst "code = " & .TextMatrix(I, 0)
        temptable.str2 = xSectionName.Caption
        temptable!date1 = DateFix(xDate.Text)
        temptable!date2 = RetLast(xCode.Text, 2)
        temptable.Str3 = TurnValue(.TextMatrix(I, 1), "", Null)
        temptable.val1 = Val(xCode.Text)
        temptable.Val10 = Val(.TextMatrix(I, 0))
       'TempTable.str4 = TurnValue(addRelationString(ItemTable!Relation, 0, .TextMatrix(I, 0)), "", Null)
        temptable.str5 = TurnValue(cnames, "", Null)
        If Not ItemTable.NoMatch Then
            If Not IsNull(ItemTable!Relation) Then
                 temptable.Str3 = temptable.Str3
            End If
        End If
        temptable.val2 = Val(.TextMatrix(I, 2))
        temptable.Val3 = Val(.TextMatrix(I, 3))
        temptable.val4 = Val(.TextMatrix(I, 4))
        temptable.val5 = Val(.TextMatrix(I, 6))
        temptable.val6 = Val(xTotal.Caption)
        temptable.val7 = Val(xLate.Text)
        temptable.val8 = Val(xOverDue.Text)
        temptable.val9 = Val(xTotal.Caption)
        temptable.Update
    Next
    End With
Next
temptable.Requery
tempdb.Execute "Insert Into Temp(Str1,Str2,Str3,Str5,Val1,Val2,Val3,Val5,Val6,val7,Val8,Val9,val10,Date1,date2)" & _
               " Select Str1,Str2,Str3,STR5,Val1,Max(Val2),Sum(val3),sum(Val5),Val6,Val7,Val8,Val9,val10,Date1,date2 FROM TEMP2" & _
               " Group by Str1,Str2,Str3,Str5,Val1,Val6,Val7,Val8,Val9,val10,Date1,date2"
myws.BeginTrans
myws.CommitTrans
Report1.Destination = nDist
Report1.WindowShowExportBtn = False
Report1.ReportFileName = MainPath & "\rpt\rpt02.rpt"
Report1.DataFiles(0) = tempFile
Report1.Action = 1
End Sub
Private Sub FillTemp()
For nTab = 0 To 3
    If Not Tab1.TabVisible(nTab) Then Exit For
    FillTab nTab
Next
End Sub
Private Sub FillTab(nTab)
Dim I As Integer
Set PRVITEMTABLE = mydb.OpenRecordset("select item,file2_10.desca,FILE2_10.[VALUE] From File2_12 inner join file2_10 on file2_12.item = file2_10.code where PaidType = " & xType.BoundText)
Set temptable = tempdb.OpenRecordset("select * from Temp")
cString = "SELECT ITEM," & _
          " FILE2_11.[value], FILE2_11.Discount " & _
          " FROM FILE2_11" & _
          " WHERE SECTION =  " & MyParn(RetMember(xCode.Text, "Section"))
Set ValueTable = mydb.OpenRecordset(cString)
nLate = CalcLateTab(nTab)
If PRVITEMTABLE.RecordCount <> 0 Then
    Do
        I = I + 1
        temptable.AddNew
        ' ≈÷«›… »‰Êœ «·«‘ —«ﬂ«   «·«Ê·Ì
        temptable!val12 = PRVITEMTABLE!Item
        temptable!str1 = PRVITEMTABLE!DESCA
        
        ValueTable.FindFirst "item = " & PRVITEMTABLE!Item
        If ValueTable.NoMatch Then
              temptable.val1 = TurnValue(ValueTable!Value, Null, 0)
              temptable.val2 = 0
              temptable.Val3 = TurnValue(ValueTable!DISCOUNT, Null, 0)
              temptable.val4 = 0
              temptable.val5 = nLate
        Else
             temptable.val1 = TurnValue(PRVITEMTABLE!Value, Null, 0)
             temptable.val2 = 0
             temptable.Val3 = 0
             temptable.val4 = 0
             temptable.val5 = nLate
        End If
    
 
       ' «÷«›… »‰Êœ «·«‘ —«ﬂ«  «·„‘ —ﬂ…
        temptable!str11 = "≈” „«—… «·≈‘ —«ﬂ ⁄‰ ⁄«„ " & paidYearString(PaidYear(xDate.Text), True)
        temptable!str12 = TurnValue(RetFind(SectionTable, "Code", "DescA", RetMember(xCode.Text, "Section")), Null, "")
        temptable!str13 = TurnValue(RetMember(xCode.Text, "DescA"), "", Null)
        temptable!str14 = MyOnly(Val(xTotal.Caption))
        temptable!val11 = TurnValue(xCode.Text, "", Null)

        temptable!date1 = xDate.Text
        temptable!VAL19 = nTab
        
        temptable!str19 = "„Ê”„ " & Tab1.TabCaption(nTab)
        temptable.val6 = Val(xLate.Text)
        temptable.val7 = Val(xOverDue.Text)
        temptable.val8 = Val(xTotal0.Caption)
        temptable.val9 = Val(xTotal.Caption)
        temptable.Val10 = Null
        temptable.Val13 = I
        temptable.Update
        PRVITEMTABLE.MoveNext
    Loop Until PRVITEMTABLE.EOF
End If
End Sub

Private Sub xType_Click(Area As Integer)
xYear.Enabled = (xType.BoundText = 1)
End Sub
Private Function RetValue(nCode, nIndex, Optional nRet = 0) As Double
If Not IsDate(xDate.Text) Then Exit Function
     
If xCode.Text = "" Then Exit Function
nYear = PaidYear(xDate.Text) - nIndex
If nYear < 1900 Then
     Exit Function
End If
    
itemSectionTable.FindFirst "CODE = " & nCode & _
" AND YEAR = " & nYear & _
" AND section = " & MyParn(RetMember(xCode.Text, "Section")) & _
" and TYPE = " & IIf(xType.BoundText = -1, 0, xType.BoundText)
If Not itemSectionTable.NoMatch Then
    RetValue = TurnValue(IIf(nRet = 0, itemSectionTable!Value, itemSectionTable!DISCOUNT), Null, 0)
    Exit Function
End If

If xType.BoundText > 0 Then
    itemSectionTable.FindFirst "CODE = " & nCode & _
    " AND YEAR = " & nYear & _
    " AND section = " & MyParn(RetMember(xCode.Text, "Section")) & _
    " and TYPE = 0"
    If Not itemSectionTable.NoMatch Then RetValue = TurnValue(IIf(nRet = 0, itemSectionTable!Value, itemSectionTable!DISCOUNT), Null, 0)
End If
End Function
Private Function addRelationString(cRelation, nTab, cItem) As String
If IsNull(cRelation) Then Exit Function
nYearPaid = PaidYear(xDate.Text) - nTab
cString = "SELECT FILE2_10.CODE, FILE2_10.DESCA, FILE2_10.ALLMEMBER, FILE2_10.LATE, FILE2_10.RELATION," & _
      " FILE2_10.ISMEMBER, FILE2_10.AGE1, FILE2_10.AGE2, FILE2_10.SEX, FILE2_10.Locker, FILE2_10.yacht, " & _
      " FILE2_10.BASICDIED,FILE2_10.BASICOLD, FILE2_10.BASICNEW, FILE2_10.MEETING, FILE2_10.DAYS, FILE2_10.NORATE, " & _
      " FILE2_40.value, FILE2_40.Discount " & _
      " FROM FILE2_10 INNER JOIN FILE2_40 ON FILE2_10.CODE = FILE2_40.item " & _
      " WHERE FILE2_40.TYPE = " & xType.BoundText & _
      " AND FILE2_40.BASIC " & _
      " AND FILE2_40.YEAR = " & nYearPaid & _
      " and FILE2_10.CODE = " & cItem & _
      " AND SECTION =  " & MyParn(RetMember(xCode.Text, "Section")) & _
      " ORDER BY FILE2_10.CODE"
Set MemItemTable = mydb.OpenRecordset(cString)
If nTab = 0 Then
    dDate = IIf(nSystem = HORSE_SYSTEM, paidDate(PaidYear(xDate.Text)), xDate.Text)
Else
    dDate = paidDate(PaidYear(xDate.Text) - nTab)
End If

Set reltable = mydb.OpenRecordset("select * From File1_11 where Member = " & xCode.Text & " and relation = " & MyParn(cRelation))
If reltable.RecordCount = 0 Then Exit Function

Do
    If Not IsNull(MemItemTable!Sex) Then
        If IsNull(reltable.Sex) Then
            GoTo endLoop
        Else
            If (MemItemTable!Sex <> reltable!Sex) Then GoTo endLoop
        End If
    End If
    
    If IsNull(MemItemTable!age1) And IsNull(MemItemTable!age2) Then
        If reltable.Relation = "2" Then
            aArray = Split(reltable.DESCA)
            cdesca = aArray(0)
        Else
            cdesca = reltable!DESCA
        End If
        addRelationString = addRelationString & IIf(addRelationString = "", "", "-") & cdesca
        GoTo endLoop
    End If
    
    If Not IsNull(MemItemTable!age1) Then
        If IsNull(reltable!DateBirth) Then GoTo endLoop
        If Age(reltable!DateBirth, dDate) < MemItemTable!age1 Then GoTo endLoop
    End If
        
    If Not IsNull(MemItemTable!age2) Then
        If IsNull(reltable!DateBirth) Then GoTo endLoop
        If Age(reltable!DateBirth, dDate) > MemItemTable!age2 Then GoTo endLoop
    End If
    If reltable.Relation = "2" Then
        aArray = Split(reltable.DESCA)
        cdesca = aArray(0)
    Else
        cdesca = reltable.DESCA
    End If
     addRelationString = addRelationString & IIf(addRelationString = "", "", "-") & cdesca
endLoop:
    reltable.MoveNext
Loop Until reltable.EOF
End Function
Private Function RelNames() As String
Dim LocalTable As Recordset
Dim cName As String
cString = "SELECT FILE2_10.CODE, FILE2_10.DESCA, FILE2_10.RELATION," & _
      " FILE2_10.ISMEMBER, FILE2_10.AGE1, FILE2_10.AGE2, FILE2_10.SEX " & _
      " FROM FILE2_10  " & _
      " WHERE  NOT ISNULL(FILE2_10.RELATION) AND SHOWPAID" & _
      " ORDER BY FILE2_10.CODE"

Set LocalTable = mydb.OpenRecordset(cString)
If LocalTable.RecordCount = 0 Then Exit Function

If nTab = 0 Then
    dDate = IIf(nSystem = HORSE_SYSTEM, paidDate(PaidYear(xDate.Text)), xDate.Text)
Else
    dDate = paidDate(PaidYear(xDate.Text) - nTab)
End If


Do
    cName = ""
    Set reltable = mydb.OpenRecordset("select * From File1_11 where Member = " & xCode.Text & " and relation = " & MyParn(LocalTable!Relation))
    If reltable.RecordCount = 0 Then GoTo EndLoop2
    Do
        If Not IsNull(LocalTable!Sex) Then
            If IsNull(reltable.Sex) Then
                GoTo endLoop
            Else
                If (LocalTable!Sex <> reltable!Sex) Then GoTo endLoop
            End If
        End If
        
        If IsNull(LocalTable!age1) And IsNull(LocalTable!age2) Then
            cName = cName & AbName(reltable!DESCA, LocalTable!Relation) & " - "
            GoTo endLoop
        End If
        
        If Not IsNull(LocalTable!age1) Then
            If IsNull(reltable!DateBirth) Then GoTo endLoop
            If Age(reltable!DateBirth, dDate) < LocalTable!age1 Then GoTo endLoop
        End If
            
        If Not IsNull(LocalTable!age2) Then
            If IsNull(reltable!DateBirth) Then GoTo endLoop
            If Age(reltable!DateBirth, dDate) > LocalTable!age2 Then GoTo endLoop
        End If
        cName = cName & AbName(reltable!DESCA, LocalTable!Relation) & " - "
endLoop:
        reltable.MoveNext
    Loop Until reltable.EOF
    If Trim(cName) <> "" Then
        cName = Mid(cName, 1, Len(cName) - 3)
        cName = Replace(LocalTable!DESCA, "“ÊÃ…", "“ÊÃ") & Space(3) & cName
        'cName = LocalTable!Desca & Space(3) & cName
        RelNames = RelNames & IIf(RelNames <> "", vbCrLf, "") & Trim(cName)
    End If
EndLoop2:
    LocalTable.MoveNext
Loop Until LocalTable.EOF
End Function
Private Function AbName(pName, cRelation)
If cRelation <> "2" Then
    AbName = Trim(pName)
    Exit Function
End If

Dim aNames
If pName = "" Then Exit Function
aNames = Split(Trim(pName))
If aNames(0) = "⁄»œ" And UBound(aNames) > 0 Then
     AbName = aNames(0) & " " & aNames(1)
Else
    AbName = aNames(0)
End If
End Function
Private Sub MYDEFINETAB()
For nTab = 0 To 3
    Grid1(nTab).rows = 1
    Grid1(nTab).AddItem ""
    If nTab > 0 Then
        LoadTabCaption
        Tab1.TabVisible(nTab) = False
    End If
Next
'LoadTabCaption
End Sub
Private Sub addMulti()
For nTab = 0 To 0
    For I = 1 To Grid1(nTab).rows - 2
        Grid1(ntabl).TextMatrix(I, 3) = nMulti
    Next
Next
Calctotals
End Sub
Private Function addMeet(Optional nRelation As String = "") As Integer
Dim loctable As New ADODB.Recordset, dBeginDate As String, dOrderDate As String

dOrderDate = GetDesca("Select DateBegin from file1_10 where code = " & xCode.Text)

If nRelation <> "" Then
    dBeginDate = GetDesca("Select DateBegin from file1_11 where Member = " & xCode.Text & " and Code = " & nRelation)
End If

cString = "Select * from file1_13 where YearCol = " & PaidYear(xDate.Text)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Do Until loctable.EOF
    cString = "Select Doc_no From file2_20 inner join file0_50  on file2_20.type = file0_50.code where Type = 0 and (not isNull([No])) and FILE2_20.code = " & xCode.Text
    cString = cString & TurnOld(cString) & "Date >= " & DateSq("2014-07-01")
    cString = cString & TurnOld(cString) & "Date <= " & DateSq("2016-02-10")
    
    'cString = cString & TurnOld(cString) & "Date >= " & DateSq(paidDate(loctable!Year))
    'cString = cString & TurnOld(cString) & "Date <= " & DateSq(loctable!Date)
    
    If GetDesca(cString) = "" Then GoTo SkipLine
        
    If dOrderDate <> "" Then
        If Age(Format(dOrderDate, "dd-mm-yyyy"), Format(loctable!Date, "dd-mm-yyyy")) < 1 Then
            GoTo SkipLine
        End If
    End If
        
    If dBeginDate <> "" Then
        If Age(Format(dBeginDate, "dd-mm-yyyy"), Format(loctable!Date, "dd-mm-yyyy")) < 1 Then
            GoTo SkipLine
        End If
    End If
        
        
    cString = "Select file1_13.code from file1_12 inner join file1_13 on file1_12.code = file1_13.code  " & _
              " where file1_13.Code = " & loctable!CODE & " and Member = " & xCode.Text
    
    If nRelation = "" Then
        cString = cString & TurnOld(cString) & " isNull(File1_12.Relation)"
    Else
        cString = cString & TurnOld(cString) & "  File1_12.relation = " & nRelation
    End If
          
    If GetDesca(cString) = "" Then
        addMeet = addMeet + 1
    End If

SkipLine:
   loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function
Private Function CheckOver() As Boolean
If nTab = 0 Then
    dDate = paidDate(PaidYear(xDate.Text))
Else
    dDate = paidDate(PaidYear(xDate.Text) - nTab)
End If
Dim cString As String
cString = "select * From File1_11"
cString = cString & turn(cString) & "Member = " & xCode.Text
cString = cString & turn(cString) & "file1_11.relation = '2'"
Set reltable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If reltable.RecordCount = 0 Then Exit Function
Dim nAge As Long
Do
    nAge = Age(reltable!DateBirth, dDate)
    If reltable!Relation = 2 And nAge >= 21 Then
        MsgBox "«Õœ «·«»‰«¡ ”‰… " & nAge
        bNoSave = True
    End If
    reltable.MoveNext
Loop Until reltable.EOF
End Function
Private Sub CompLookup()
Dim Generalarray(5)
Set Generalarray(0) = Me
Dim cString As String
Dim listarray(0, 4)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
nYear = PaidYear(xDate.Text) - 1
cString = "Select file1_10.code,Desca,Max(Format(File2_20.date,'yyyy/mm/dd')) as Lastdate From file1_10 left join file2_20 on file1_10.code = file2_20.code "
cString = cString & turn(cString) & "ISSAVE = FALSE"
cString = cString & turn(cString) & "DateLast >= " & DateSq(paidDate(nYear))
cString = cString & turn(cString) & "DateLast <= " & DateSq(paidDate(nYear, True))
If xCompany.MatchedWithList Then cString = cString & turn(cString) & "file1_10.company = " & MyParn(xCompany.BoundText)

Generalarray(1) = cString
Generalarray(2) = "Group by file1_10.code,Desca  HAVING Max(file2_20.date)  < " & DateSq(paidDate(nYear + 1)) & " Order by FILE1_10.CODE"
Generalarray(3) = 4200
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%DescA%%) "

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 3000

GrdArray(2, 0) = " «—ÌŒ «Œ— «” „«—…"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchComp.bEnter = True
oSearchComp.Caption = "«” ⁄·«„ «” „«—«  «·‘—ﬂ« "
oSearchComp.Show 1
End Sub
Private Sub CompLookup2(Optional cFilter As String = "")
Dim Generalarray(5)
Set Generalarray(0) = Me
Dim cString As String
Dim listarray(2, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me

nYear = PaidYear(xDate.Text)
cString = "Select File2_20.Doc_No,File1_10.Code,file1_10.DescA,[No],Format([Date],'yyyy/m/d') From File2_20 Inner Join File1_10 On  File2_20.Code =  File1_10.Code"
cString = cString & turn(cString) & "ISSAVE = FALSE"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & turn(cString) & "file2_20.Date >= " & DateSq(paidDate(nYear))
cString = cString & turn(cString) & "file2_20.Date <= " & DateSq(paidDate(nYear, True))
If xCompany.MatchedWithList Then cString = cString & turn(cString) & "file1_10.company = " & MyParn(xCompany.BoundText)

Generalarray(1) = cString
Generalarray(2) = "group by File2_20.Doc_No,File1_10.Code,file1_10.DescA, [No],[Date] order by  file2_20.doc_no"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·—ﬁ„ «Ê «·«”„"
listarray(0, 1) = "(FILE1_10.Code Like val('cFilter') or %%FILE1_10.DescA%%)"

If cFilter = "" Then
    listarray(1, 0) = "—ﬁ„ «·ﬁ”Ì„… «Ê —ﬁ„ «·„” ‰œ"
    listarray(1, 1) = "( VAL([NO] & '') = VAL('cFilter') OR Val(File2_20.Doc_No) Like Val('cFilter') )"
Else
    listarray(1, 0) = "—ﬁ„ «·„” ‰œ"
    listarray(1, 1) = "(Val(File2_20.Doc_No) Like Val('cFilter') )"
End If

listarray(2, 0) = "«· «—ÌŒ √Ê »«·„Ê”„"
listarray(2, 1) = "(##Date##  OR VAL(FILE2_20.[YEAR] & '') = VAL('cFilter') )"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·⁄÷Ê"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "≈”„ «·⁄÷Ê"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "—ﬁ„ «·«” „«—…"
GrdArray(3, 1) = IIf(cFilter = "", 1000, 0)

GrdArray(4, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(4, 1) = 1600

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchComp.bEnter = False
oSearchComp.Caption = "«” ⁄·«„ «” „«—«  «·‘—ﬂ« "
oSearchComp.Show 1
End Sub
