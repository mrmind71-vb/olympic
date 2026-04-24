VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form items1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "»Ì«‰«  «·√’‰«›"
   ClientHeight    =   6300
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   9570
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
   ScaleHeight     =   6300
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox xSerial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "·Â —ﬁ„ „”·”·"
      Height          =   240
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   2700
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   825
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.TextBox xItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6600
      MaxLength       =   15
      TabIndex        =   0
      Top             =   750
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9570
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5760
      Width           =   9570
      Begin VB.CommandButton CmdPrevious 
         BackColor       =   &H00C0FFFF&
         Caption         =   "”«»ﬁ"
         Height          =   390
         Left            =   6900
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   75
         Width           =   1290
      End
      Begin VB.CommandButton CmdNext 
         BackColor       =   &H00C0FFFF&
         Caption         =   "·«Õﬁ"
         Height          =   390
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   75
         Width           =   1215
      End
      Begin VB.CommandButton CmdFirst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√Ê·"
         Height          =   390
         Left            =   4575
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   75
         Width           =   1140
      End
      Begin VB.CommandButton CmdLast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "√ŒÌ—"
         Height          =   390
         Left            =   3525
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   75
         Width           =   1065
      End
      Begin VB.Label xRecordNumber 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   150
         Width           =   3165
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9570
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   9570
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1725
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0FFFF&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3555
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5385
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   150
         Width           =   915
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4470
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   150
      RightToLeft     =   -1  'True
      ScaleHeight     =   2385
      ScaleWidth      =   8910
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3225
      Width           =   8940
      Begin VB.Label xInPut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1500
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê«—œ"
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1578
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "’«œ—"
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
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ï „‘ —Ì«  "
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
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   686
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ï „—œÊœ „‘ —Ì« "
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
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1132
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ï Â«·ﬂ"
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2025
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ï „—œÊœ „»Ì⁄«  "
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1132
         Width           =   1710
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ï „»Ì⁄« "
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   686
         Width           =   1155
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ Õ«·Ï "
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   885
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   150
         Width           =   915
      End
      Begin VB.Label xSales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   915
      End
      Begin VB.Label xRetSales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label xDamage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label xPurchase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   525
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label xRetPurchase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   525
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label xOutPut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   525
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1500
         Width           =   1215
      End
   End
   Begin VB.TextBox xMin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6675
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1875
      Width           =   1140
   End
   Begin VB.TextBox xUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1875
      Width           =   1215
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4275
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   3540
   End
   Begin VB.TextBox xPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1875
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "file_2"
      Top             =   1350
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox xCost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6675
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2250
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo xGroup 
      Bindings        =   "Items1.frx":0000
      DataSource      =   "Data2"
      Height          =   315
      Left            =   4275
      TabIndex        =   2
      Top             =   1500
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   -2147483643
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰ «·’‰›"
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
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄…"
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
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1575
      Width           =   750
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”⁄— »Ì⁄"
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
      Left            =   3225
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2325
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊÕœ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3225
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1950
      Width           =   510
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Õœ ≈⁄«œ… «·ÿ·»"
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
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1950
      Width           =   1185
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”⁄—  ﬂ·›…"
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
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2325
      Width           =   825
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·’‰› "
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
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   825
      Width           =   900
   End
End
Attribute VB_Name = "items1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte, GroupTable As Recordset, CodeTable As Recordset
Dim cardtable As Recordset, movetable As Recordset, DetailTable As Recordset, nRecordNumber As Integer
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols EmptyMode
myDefine
MsgBox "·«  ÊÃœ ”Ã·«  „‰ ›÷·ﬂ «÷€ÿ ⁄·Ì “— «·«÷«›… ·«÷«›… ”Ã·«  ÃœÌœ…"
End Sub
Sub AddProc()
If Not formMode = Editmode Then
    Handlecontrols addmode
End If
    myDefine
    Load ItemCode
    ItemCode.Show 1
End Sub
Sub Handlecontrols(nMode)
Select Case nMode
Case Editmode
     cmdAdd.Enabled = True
     CmdDel.Enabled = True
     CmdInform.Enabled = True
     CmdExit.Enabled = True
     CmdSave.Enabled = True
     CmdUndo.Enabled = True
     CmdPrevious.Enabled = True
     CmdNext.Enabled = True
     CmdLast.Enabled = True
     CmdFirst.Enabled = True
     xItem.Enabled = False
Case addmode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    cmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xItem.Enabled = True
    xItem.SetFocus
Case EmptyMode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xItem.Enabled = True
End Select
End Sub
Sub ItemsLookup()
    Dim Generalarray(4)
    Dim GrdArray(2)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Item as «·’‰›,DescA as [«”„ «·’‰›] From file1_10 "
    Generalarray(3) = " Where DescA Like('*cFilter*')"
    Generalarray(4) = "Order by Item"
    
    GrdArray(1) = 1500
    GrdArray(2) = 4500
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols Editmode
End Sub
Sub myDefine()
xItem.Text = ""
xDescA.Text = ""
xUnit.Text = ""
xGroup.BoundText = ""
xCost.Text = ""
xPrice.Text = ""
xMin.Text = ""
'xSerial.Value = 0
'xEndPrice.Text = ""
'xList.BoundText = "1"
xBalance.Caption = ""
xSales.Caption = ""
xRetSales.Caption = ""
xPurchase.Caption = ""
xRetPurchase.Caption = ""
xDamage.Caption = ""
xOutPut.Caption = ""
'xIntPut.Caption = ""
xRecordNumber.Caption = ""
End Sub
Sub myProc()
If formMode = Editmode Then
   cardtable.FindFirst "item = " & MyParn(GrdText(Search.Grid1, 0))
   MyLoad
End If
End Sub
Sub MyLoad()
xItem.Text = cardtable.Item
xDescA.Text = cardtable.DESCA
xUnit.Text = TurnValue(cardtable.Unit, Null, "")
xMin = Format(cardtable!Min)
xGroup.BoundText = TurnValue(cardtable.Group, Null, "")
xCost.Text = TurnValue(Format(cardtable.COST, "##0.00"), Null, "")
xPrice.Text = TurnValue(Format(cardtable.price, "##0.00"), Null, "")
'xSerial.Value = IIf(cardtable.serial, 1, 0)
DetailTable.FindFirst "item = " & MyParn(xItem.Text)
If Not DetailTable.NoMatch Then
    xBalance.Caption = Format(DetailTable.balance, "##0.00")
    xSales.Caption = Format(DetailTable.Sales, "##0.00")
    xRetSales.Caption = Format(DetailTable.RetSales, "##0.00")
    xPurchase.Caption = Format(DetailTable.Purchase, "##0.00")
    xRetPurchase.Caption = Format(DetailTable.RetPurchase, "##0.00")
    xDamage.Caption = Format(DetailTable.damage, "##0.00")
    xOutPut.Caption = Format(DetailTable.Output, "##0.00")
    xInPut.Caption = Format(DetailTable.Input, "##0.00")
Else
    xBalance.Caption = Format(0, "##0.00")
    xSales.Caption = Format(0, "##0.00")
    xRetSales.Caption = Format(0, "##0.00")
    xPurchase.Caption = Format(0, "##0.00")
    xRetPurchase.Caption = Format(0, "##0.00")
    xDamage.Caption = Format(0, "##0.00")
    xOutPut.Caption = Format(0, "##0.00")
    xInPut.Caption = Format(0, "##0.00")
End If
xRecordNumber = "”Ã· " & cardtable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
End Sub
Sub MyReplace()
cardtable.FindFirst "item = " & MyParn(xItem.Text)
If cardtable.NoMatch Then
    cardtable.AddNew
    formMode = addmode
Else
    cardtable.Edit
    formMode = Editmode
End If
cardtable.Item = xItem.Text
cardtable.DESCA = xDescA.Text
cardtable.Unit = TurnValue(xUnit.Text, "", Null)
cardtable.Group = TurnValue(xGroup.BoundText, "", Null)
cardtable.COST = Val(xCost.Text)
cardtable.price = Val(xPrice.Text)
cardtable!Min = Val(xMin.Text)
'cardtable!serial = IIf(xSerial.Value = 0, False, True)
cardtable.Update
End Sub
Function MYVALID() As Boolean
If xItem.Text = "" Then
    MsgBox "ﬂÊœ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If formMode <> Editmode Then
    cardtable.FindFirst " item = " & MyParn(xItem.Text)
    If Not cardtable.NoMatch Then
        MsgBox "Â–« «·’‰› „ÊÃÊœ „‰ ﬁ»·"
        Exit Function
    End If
End If

If xDescA.Text = "" Then
    MsgBox "·«  ÊÃœ »Ì«‰« "
    Exit Function
End If
        
If xGroup.BoundText = "" Then
    MsgBox "·«»œ „‰ «œ—«Ã „Ã„Ê⁄… ··’‰›"
    Exit Function
End If

MYVALID = True
End Function
Private Sub CmdAdd_Click()
    AddProc
End Sub
Private Sub CmdDel_Click()
If Not myValidDelete Then Exit Sub
On Error GoTo MyError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
        cardtable.Delete
        cardtable.Requery
        nRecordNumber = cardtable.RecordCount
        If cardtable.RecordCount > 0 Then
            cardtable.MoveLast
            MyLoad
        Else
            EmptyProc
        End If
End If
Exit Sub
MyError:
If Err.Number = 3200 Then MsgBox "«·„·› ·Â Õ—ﬂ…"
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
cardtable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
ItemsLookup
End Sub
Private Sub CmdLast_Click()
cardtable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
cardtable.MoveNext
If cardtable.EOF Then
    cardtable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
cardtable.MovePrevious
If cardtable.BOF Then
    cardtable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub cmdSave_Click()
msgBoxStr = "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø"
If Not MYVALID Then Exit Sub
SendKeys "{tab}"
MyReplace
Select Case formMode
Case addmode, EmptyMode
    AddProc
    cardtable.Requery
Case Editmode
    editProc
End Select
MyMove = True
End Sub
Private Sub CmdUndo_Click()
Select Case formMode
Case EmptyMode
    myDefine
Case addmode
    If cardtable.RecordCount > 0 Then cardtable.MoveLast
    nRecordNumber = cardtable.RecordCount

    formMode = Editmode
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo) Then SendKeys "{tAB}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If cardtable.RecordCount = 0 Then Exit Sub
If formMode <> Editmode Then Exit Sub
If KeyCode = 34 Then CmdPrevious_Click
If KeyCode = 33 Then CmdNext_Click
End Sub
Private Sub Form_Load()
Set cardtable = mydb.OpenRecordset("Select * from File1_10  order by ITEM ", dbOpenDynaset)
Set movetable = mydb.OpenRecordset("file1_11", dbOpenDynaset)
Set GroupTable = mydb.OpenRecordset("file1_50", dbOpenSnapshot)
Set CodeTable = mydb.OpenRecordset("file1_70", dbOpenDynaset)
Set DetailTable = mydb.OpenRecordset(DetailString, dbOpenSnapshot)
     
Data2.DatabaseName = MdbPath
Data2.RecordSource = "FILE1_50"
Data2.Refresh
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
nRecordNumber = cardtable.RecordCount
If cardtable.RecordCount > 0 Then
    cardtable.MoveLast
    editProc
    MyLoad
Else
    EmptyProc
End If
End Sub
Function DetailString()
DetailString = "Select Item," & _
"Sum(iif(isNull([In]),0,[In])-iif(isNull(Out),0,Out)) As Balance," & _
myiif("[Type] = '6' ", "Out ") & " As Sales," & _
myiif("[Type] = '3' ", "[In] ") & " As RetSales," & _
myiif("[Type] = '2' ", "[In] ") & " As Purchase, " & _
myiif("[Type] = '7' ", "Out") & " As RetPurchase, " & _
myiif("[Type] = '9' ", "Out") & " As Damage, " & _
myiif("[Type] = '4' ", "[In]") & " As Input," & _
myiif("[Type] = '8' ", "Out") & " As OutPut " & _
" From file1_11 " & _
" Group by item "
End Function
Private Sub Fiximage(cImage)
Image1.Top = 0
Image1.Left = 0
Image1.Picture = LoadPicture(cImage)
Image1.Width = Me.Width
Image1.Height = Me.Height - 500
End Sub
Private Function myValidDelete() As Boolean
movetable.FindFirst " item = " & MyParn(xItem.Text)
If Not movetable.NoMatch Then
    MsgBox "«·’‰› ·œÌÂ Õ—ﬂ… ·« Ì„ﬂ‰ «·€«ƒÂ"
    Exit Function
End If
myValidDelete = True
End Function
Private Sub Text1_LostFocus()
Text1.Text = RetAbbr(Text1.Text)
End Sub
Private Function RetAbbr(pString) As String
Dim aString
If pString = "" Then Exit Function
aString = Split(pString)
For I = 0 To UBound(aString)
    cString = RetCondition2(UBound(aString) - I, aString)
    cardtable.FindFirst cString
    If Not cardtable.NoMatch Then Exit For
Next
RetAbbr = IIf(cardtable.NoMatch, "", cardtable.DESCA)
End Function
Private Function RetCondition2(nCount, aString)
For I = 0 To nCount
    If I = 0 Then
        cString = "Desca Like " & MyParn(aString(I) & "*") & " and "
    Else
        cString = cString & "Instr(1,DescA," & MyParn(" " & aString(I)) & ") > 0" & " and "
    End If
Next
'For i = 0 To nCount
'    cString = cString & "Instr(descA) " & MyParn(IIf(i = 0, "", " ") & aString(i) & "*") & " and "
'Next
RetCondition2 = Left(cString, Len(cString) - 4)
End Function
Private Function RetCondition(nCount, aString)
Dim cString As String
cString = "DescA Like "
For I = 0 To nCount
  cString = cString + MyParn(aString(I) & "*") & " + "
Next
'For i = 0 To nCount
'    cString = cString & "Instr(descA) " & MyParn(IIf(i = 0, "", " ") & aString(i) & "*") & " and "
'Next
RetCondition = Left(cString, Len(cString) - 3)
End Function
