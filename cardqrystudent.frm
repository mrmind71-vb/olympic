VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cardqryStudentfrm 
   Caption         =   "ÿ»«⁄… «·ﬂ«—‰ÌÂ« "
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18660
   FillColor       =   &H000000C0&
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   18660
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   13815
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -45
      Width           =   4605
      Begin VB.CommandButton cmdLastFillgrd 
         Caption         =   "«” —Ã«⁄ «Œ— ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton cmdSavePrint 
         Caption         =   " „  «·ÿ»«⁄…"
         Height          =   390
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqrystudent.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqrystudent.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.Frame frmProg1 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5445
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   8145
      Width           =   5100
      Begin ComctlLib.ProgressBar prog1 
         Height          =   555
         Left            =   45
         TabIndex        =   41
         Top             =   180
         Visible         =   0   'False
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   979
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   13635
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8145
      Width           =   2310
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "«·’› :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "«·⁄„Êœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "÷»ÿ «·ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   15975
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8145
      Width           =   2490
      Begin VB.TextBox xRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   315
         Width           =   435
      End
      Begin VB.TextBox xDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label8 
         Caption         =   "«”›· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label9 
         Caption         =   "Ì„Ì‰ :"
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
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   10575
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8145
      Width           =   3030
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " „  ÿ»«⁄ Â"
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
         Index           =   0
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   315
         Width           =   1005
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   0
         Left            =   1170
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»œÊ‰ ’Ê—…"
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
         Index           =   1
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   315
         Width           =   915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Index           =   1
         Left            =   2655
         Shape           =   5  'Rounded Square
         Top             =   315
         Width           =   240
      End
   End
   Begin VB.Frame Frame5 
      Height          =   780
      Left            =   5445
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1215
      Width           =   6855
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "cardqrystudent.frx":4D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   3420
         Picture         =   "cardqrystudent.frx":74F1
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1680
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   555
         Left            =   5130
         TabIndex        =   4
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cardqrystudent.frx":991B
         Caption         =   "√÷«›… "
         Alignment       =   4
      End
      Begin Threed.SSCommand cmdprintrep 
         Height          =   555
         Left            =   1710
         TabIndex        =   15
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cardqrystudent.frx":C984
         Caption         =   "ÿ»«⁄…  ﬁ—Ì—"
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "„Ê”„ «·ÿ»«⁄…"
      Height          =   780
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1215
      Width           =   1905
      Begin VB.TextBox xSeason 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1590
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   " Õﬁﬁ „‰ «·ﬂ«—‰Ì…"
      Height          =   1230
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   765
      Width           =   2760
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   9
         Top             =   315
         Width           =   2580
      End
      Begin VB.Label xUnCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001111AE&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   765
         Width           =   2580
      End
   End
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·–Ì ·„ Ìÿ»⁄ ›ﬁÿ"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9990
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   900
      Width           =   2265
   End
   Begin VB.CheckBox xPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·„”œœÌ‰ ›ﬁÿ"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   900
      Width           =   2265
   End
   Begin VB.CheckBox xDamage 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… »œ· ›«ﬁœ Ê»œ·  «·›"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   2265
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   9855
      Width           =   18660
      _ExtentX        =   32914
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   420
      Left            =   990
      Top             =   855
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   375
      Left            =   0
      Top             =   810
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
      Height          =   375
      Left            =   135
      Top             =   135
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   375
      Left            =   2430
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
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
      BackColor       =   128
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   420
      Left            =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   420
      Left            =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
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
      Height          =   1365
      Left            =   12330
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   630
      Width           =   6135
      Begin VB.TextBox xCode1 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   960
      End
      Begin VB.TextBox xCode2 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox xDate1 
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox xDate2 
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   945
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "«·ﬂÊœ"
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   690
      End
      Begin VB.Label xcode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   28
         Top             =   225
         Width           =   3840
      End
      Begin VB.Label Label10 
         Caption         =   "≈·Ì"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   1005
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "cardqrystudent.frx":EF88
      Height          =   6000
      Left            =   90
      TabIndex        =   44
      Top             =   2025
      Width           =   18375
      _cx             =   32411
      _cy             =   10583
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      TabBehavior     =   0
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
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   8010
      Width           =   2760
      Begin VB.CommandButton Command1 
         Caption         =   "«Œ Ì«— «·ﬂ·"
         Height          =   555
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   2625
      End
   End
End
Attribute VB_Name = "cardqryStudentfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileSave As String, cFileSave2 As String, cFilePrint As String, aDegree_group As Variant
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim printTable As New ADODB.Recordset
Private Sub cmdAdd_Click()
'checkErr
myloadGrd
cmdPrint.Enabled = (grid1.rows > 1)
checkPhoto
On Error Resume Next
grid1.SaveGrid cFileSave, flexFileData
Err.Clear
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·ﬂ· !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
    grid1.rows = 1
    grid1.SaveGrid cFileSave, flexFileData
'    DefineText Me
    CalcTotals
End If
End Sub

Private Sub cmdExel_Click()
For I = 1 To grid1.rows - 1
    If Not validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) Then
        grid1.RowHidden(I) = True
    End If
Next
ToFileExel grid1
For I = 1 To grid1.rows - 1
    grid1.RowHidden(I) = False
Next
End Sub

Private Sub CmdPrint_Click()
If grid1.rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
    Exit Sub
End If

If Val(xRow.Text) > 5 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 2 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
If Not doprint Then
    MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
    Exit Sub
End If
Set CardPrintNew.myForm = Me
CardPrintNew.PrintArray
CardPrintNew.Show 1
SavePrinted

grid1.SaveGrid cFileSave2, flexFileData
For I = grid1.rows - 1 To 1 Step -1
    If Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0 Then
        grid1.RemoveItem I
    End If
Next
grid1.SaveGrid cFileSave, flexFileData
checkPhoto
'End If
End Sub
Private Sub CmdExit_Click()
Unload Me
Set cardqryStudentfrm = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdClear_Click()
grid1.rows = 1
End Sub
Private Sub cmdMember_Click()
Member.Show 1
End Sub
Private Sub CmdLastFillGrd_Click()
Dim fs As New FileSystemObject
If fs.FileExists(cFileSave) Then
    grid1.LoadGrid cFileSave, flexFileData
    If grid1.rows > 1 Then cmdPrint.Enabled = True
    checkPhoto
End If
End Sub

Private Sub cmdprintrep_Click()
Set PrintGrdNew.myForm = Me
Dim I As Long
For I = 1 To grid1.rows - 1
    If Not validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) Then
        grid1.RowHidden(I) = True
    End If
Next
grid1.ColHidden(grid1.Cols - 1) = True
PrintGrdNew.doprint grid1, 0.8, -3, "ÿ»«⁄… «·ÿ·»…", , , , False, False, 9
PrintGrdNew.Show 1
grid1.ColHidden(grid1.Cols - 1) = False
For I = 1 To grid1.rows - 1
    grid1.RowHidden(I) = False
Next
End Sub

Private Sub Command1_Click()
For I = 1 To grid1.rows - 1
    grid1.TextMatrix(I, grid1.Cols - 1) = 1
Next
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1.LoadGrid cFileSave2, flexFileData
    For I = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(I, grid1.Cols - 1)) = 0 Then
            grid1.RemoveItem I
        End If
    Next
    cmdPrint.Enabled = (grid1.rows > 1)
    checkPhoto
End If
End Sub
Private Sub Form_Load()
openCon con
cFileSave = App.Path & "\" & Me.Name & ".grd"
cFileSave2 = App.Path & "\" & Me.Name & "_printed.grd"
Fixgrd
LoadText Me
xSeason.Text = sSeason_Student
xPrinted.Value = 1
xPaid.Value = 1
If retFlag(aSec, "DAMAGE") Then xPrinted.Enabled = True
grid1.ExplorerBar = flexExSortAndMove
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
grid1.SaveGrid cFileSave, flexFileData
End Sub

Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
grid1.SaveGrid cFileSave, flexFileData
End Sub
Private Sub grid1_EnterCell()
grid1.Editable = IIf(grid1.Col = grid1.Cols - 1, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
With grid1
If .rows = 1 Then Exit Sub
If KeyCode = 46 Then
    .RemoveItem grid1.Row
    .SaveGrid cFileSave, flexFileData
    CalcTotals
End If
End With
End Sub
Private Sub xCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    xUnCode.Caption = ""
    If xCode.Text = "" Then Exit Sub
    xUnCode.ForeColor = -2147483630
    If Val(unMyCodeBar(xCode.Text, "3")) <> 1 Then
        xUnCode.Caption = "Error"
        xUnCode.ForeColor = vbRed
    Else
        xUnCode.Caption = unMyCodeBar(xCode.Text)
    End If
    myGotFocus xCode
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    'StudentLookupAll Me, oSearch
End If
End Sub

Private Sub xCode1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub xCode1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim cWhere As String
    'StudentLookupAll Me, oSearch
End If
End Sub

Private Sub xCode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub

Private Sub xCODE2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    WorkerLookupAll Me, oSearch
End If
End Sub
Private Function CountGrid() As Integer
With grid1
For I = 1 To grid1.rows - 1
    'If .TextMatrix(I, 6) = True Then CountGrid = CountGrid + 1
    CountGrid = CountGrid + 1
Next
End With
End Function
Private Sub countPrint()
nCountPrint = 0
With grid1
For I = 1 To .rows - 1
   If .TextMatrix(I, 6) = True Then nCountPrint = nCountPrint + 1
Next
lblCount.Caption = nCountPrint / 1
End With
End Sub
Private Function MakeString()
MakeString = "#" & ";"
MakeString = MakeString & "|#" & 0 & ";" & "ﬂ«—‰ÌÂ ÃœÌœ"
MakeString = MakeString & "|#" & 1 & ";" & "»œ· ›«ﬁœ"
End Function
Private Sub SavePrinted()
With grid1
dTime = Time
dDate = Date
Dim aInsert As Variant
con.BeginTrans
For I = 1 To .rows - 1
   If validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) And (Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0) Then
        aInsert = AddFlag(Empty, "CODE", addstring(grid1.TextMatrix(I, 0)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(I, 1)))
        
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(Now, "YYYY-MM-DD HH:NN")))
        aInsert = AddFlag(aInsert, "TYPE", addvalue(grid1.TextMatrix(I, 7)))
        aInsert = AddFlag(aInsert, "YEAR", addvalue(sSeason_Student))
        aInsert = AddFlag(aInsert, "DESCA_TYPE", addstring(grid1.TextMatrix(I, 8)))
        aInsert = AddFlag(aInsert, "[GROUP]", addvalue(grid1.TextMatrix(I, 9)))
        con.Execute addInsert(aInsert, "FILE4_30")
        aInsert = AddFlag(Empty, "DATE_PRINT", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        con.Execute addUpdate(aInsert, "FILE3_10", "FILE3_10.CODE = " & grid1.TextMatrix(I, 0))
    End If
Next
con.CommitTrans
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
End Sub
Function eofGrd(cId) As Boolean
eofGrd = (grid1.FindRow(cId, , 0) = -1)
End Function
Private Function doprint() As Boolean
SettingArray(cUpMargin) = MyMeasure(1.9) + MyMeasure(Val(xDown.Text) / 10)
SettingArray(cRightMargin) = MyMeasure(1.4) + MyMeasure(Val(xRight.Text) / 10)
SettingArray(cCardWidth) = MyMeasure(9.65)
SettingArray(cCardHeight) = MyMeasure(5.81)
SettingArray(cRows) = 5
SettingArray(cCols) = 2
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nspace = 0.62
nup = 0.3

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * nCols) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For I = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡


prog1.Value = 0
prog1.Visible = True

For I = 1 To grid1.rows - 1
    If validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) And (Val(grid1.TextMatrix(I, grid1.Cols - 1)) <> 0) Then
        If grid1.TextMatrix(I, 9) = "1" Then
            nspace = 0.6
        Else
            nspace = 0.6
        End If

        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        nDiffer = 1.1
        nLine = 1
        
        ' „Â‰œ”
        tCard.AddNew
        tCard!Right = MyMeasure(0)
        tCard!Top = MyMeasure(1.15) - MyMeasure(nDiffer) - MyMeasure(0.5) + MyMeasure(0.1)
        tCard!Width = MyMeasure(9)
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = &H80&
        tCard!FontSize = 13
        tCard!Text = grid1.TextMatrix(I, 10)
        tCard!TextAlign = taCenterTop
        tCard!CardNo = nCard
        tCard.Update
        
        ' „Â‰œ”
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.2) + MyMeasure(0.3) + MyMeasure(0.55)
        tCard!Top = MyMeasure(1.15) - MyMeasure(nDiffer) + MyMeasure(0.2)
        tCard!Width = MyMeasure(1.5)
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = &HC0&
        tCard!FontSize = 11
        tCard!Text = ": «·«”‹‹„"
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        
                
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.2) + MyMeasure(0.3) + MyMeasure(0.3)
        tCard!Top = MyMeasure(1.25) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.4)
        tCard!Width = MyMeasure(5.5)
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!FontSize = 11
        tCard!Text = TurnValue(grid1.TextMatrix(I, 1), "", Null)
        tCard!CardNo = nCard
        tCard.Update
        
        ' —ﬁ„ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
        tCard!Width = MyMeasure(2.1)
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = &HC0&
        tCard!FontSize = 11
        tCard!Text = ": ⁄÷ÊÌ‹… —ﬁ„"
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
        
        tCard.AddNew
        tCard!Right = MyMeasure(1.25) + MyMeasure(nDiffer) + MyMeasure(0.3) + MyMeasure(0.3)
        tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
        tCard!Width = 0
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = vbBlack
        tCard!FontSize = 11
        tCard!Text = ArbString(grid1.TextMatrix(I, 0))
        tCard!TextAlign = taLeftTop
        tCard!CardNo = nCard
        tCard.Update
       
'    .TextMatrix(0, 0) = "—ﬁ„ «·⁄÷Ê"
'    .TextMatrix(0, 1) = "«·«”„"
'    .TextMatrix(0, 2) = "«·›—ﬁ…"
'    .TextMatrix(0, 3) = "«·‘⁄»…"
'    .TextMatrix(0, 4) = "Ì‰ ÂÏ ›Ï"
'
'    .TextMatrix(0, 5) = "—ﬁ„ «·«Ì’«·"
'    .TextMatrix(0, 6) = " «—ÌŒ «·ÿ»«⁄…"
'    .TextMatrix(0, 7) = "‰Ê⁄ «·ﬂ«—‰ÌÂ"
'    .TextMatrix(0, 8) = "‰Ê⁄ «·ﬂ«—‰ÌÂ"
'    .TextMatrix(0, 9) = "«·„Ã„Ê⁄…"
'    .TextMatrix(0, 10) = "«Œ Ì«—"
       
        ' «·„—«›ﬁ
        If grid1.TextMatrix(I, 9) = "1" Then
            nLine = nLine + 1
            tCard.AddNew
            tCard!Right = MyMeasure(0.5) + MyMeasure(0.1) + MyMeasure(0.3)
            tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
            tCard!Width = MyMeasure(2.8)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = &HC0&
            tCard!FontSize = 11
            tCard!Text = ": «·›—ﬁ…"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(1.85) + MyMeasure(0.3)
            tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
            tCard!Width = MyMeasure(5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!Text = TurnValue(grid1.TextMatrix(I, 2) & turn(grid1.TextMatrix(I, 3), " - ") & grid1.TextMatrix(I, 3))
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
        Else
            
            nLine = nLine + 1
            ' ’·… «·ﬁ—«»…
            tCard.AddNew
            tCard!Right = MyMeasure(0.5) + MyMeasure(0.1) + MyMeasure(0.3)
            tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
            tCard!Width = MyMeasure(2.1)
            
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = &HC0&
            tCard!FontSize = 12
            tCard!Text = " : «·‘⁄»‹‹…"
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Right = MyMeasure(2) + MyMeasure(0.25) + MyMeasure(0.3)
            tCard!Top = MyMeasure(1.15) + MyMeasure(nspace * nLine) - MyMeasure(nDiffer) + MyMeasure(0.2) + MyMeasure(0.6)
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = vbBlack
            tCard!FontSize = 11
            tCard!Text = TurnValue(grid1.TextMatrix(I, 3), "", Null)
            tCard!TextAlign = taRightTop
            tCard!CardNo = nCard
            tCard.Update
        End If
        nLine = nLine + 1
        ' «·ﬂ«—‰ÌÂ ‘Œ’Ì
        tCard.AddNew
        tCard!Right = MyMeasure(0.5) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Top = MyMeasure(2.4) + MyMeasure(0.2)
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = vbRed
        tCard!FontSize = 8
        'tCard!Text = "Â–« «·ﬂ«—‰ÌÂ ‘Œ’Ì Ìﬁœ„ ⁄‰œ ﬂ· ÿ·»"
        tCard!Width = MyMeasure(4.8)
        tCard!Height = MyMeasure(0.6)
        tCard!ISBARCODE = True
        tCard!Text = MyCodeBar(grid1.TextMatrix(I, 0), "3")
        tCard!TextAlign = taRightTop
        tCard!CardNo = nCard
        tCard.Update
       
'        ' «·ﬂ«—‰ÌÂ ‘Œ’Ì
'        tCard.AddNew
'        tCard!Right = MyMeasure(0.8) + MyMeasure(0.8)
'        tCard!Top = MyMeasure(1.5) + MyMeasure(nSpace * nLine) + MyMeasure(0.1) - MyMeasure(nDiffer)
'        tCard!Width = MyMeasure(3.8)
'        tCard!Height = 0
'        tCard!FontName = "Arial"
'        tCard!FontBold = True
'        tCard!ForeColor = vbRed
'        tCard!fontsize = 8
'        'tCard!Text = "ÊÌ”Õ» ›Ï Õ«·… «⁄«— Â ··€Ì—"
'        tCard!TextAlign = taCenterTop
'        tCard!CardNo = nCard
'        tCard.Update
       
'        tCard.AddNew
'        tCard!Right = MyMeasure(0.5)
'        tCard!Top = MyMeasure(1.1) + MyMeasure(0.65 * 4) + MyMeasure(0.4) - MyMeasure(0.1)
'        tCard!Width = 0
'        tCard!Height = 0
'        tCard!FontName = "Arial"
'        tCard!FontBold = True
'        tCard!ForeColor = &HFF&
'        tCard!FontSize = 10
'        tCard!Text = "Ê·« Ì”„Õ »≈⁄«— Â ··€Ì—"
'        tCard!TextAlign = taRightTop
'        tCard!CardNo = nCard
'        tCard.Update
       
       
        ' Ì‰ ÂÌ ›Ì
        tCard.AddNew
        tCard!Right = MyMeasure(2.5) + MyMeasure(0.3)
        tCard!Top = MyMeasure(3.1) + MyMeasure(0.2)
        tCard!Width = 0
        tCard!Height = 0
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = vbRed
        tCard!FontSize = 9
        tCard!Text = "Ì‰ ÂÌ ›Ì " & Format(grid1.TextMatrix(I, 4), "yyyy/m/d")
        tCard!CardNo = nCard
        tCard.Update
        
' «·Ã“¡ «·«Ì”—
        
        'ﬂ·„… —∆Ì” ‰«œÌ «·«”ﬂ‰œ—Ì…
        tCard.AddNew
        tCard!Right = MyMeasure(6) - MyMeasure(0.3)
        tCard!Top = MyMeasure(2.9) - MyMeasure(0.1) - MyMeasure(nDiffer) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Width = MyMeasure(3)
        tCard!Height = 1000
        tCard!FontName = "Arial"
        tCard!FontBold = True
        tCard!ForeColor = &HFF&
        tCard!FontSize = 9
        tCard!TextAlign = taCenterTop
        tCard!Text = "‰ﬁÌ» „Â‰œ”Ì «·«”ﬂ‰œ—Ì…"
        tCard!CardNo = nCard
        tCard.Update
                
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.4) - MyMeasure(0.2) + MyMeasure(0.3) - MyMeasure(0.3)
        tCard!Top = MyMeasure(0.62) - MyMeasure(nDiffer) + MyMeasure(0.3)
        tCard!Width = MyMeasure(2.4) * 0.8
        tCard!Height = MyMeasure(2.8) * 0.8
        tCard!Text = RetPhoto_s(grid1.TextMatrix(I, 0))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        
'        If validPhoto(RetPhoto_s(grid1.TextMatrix(I, 7)) & "") Then
'            '«·’Ê—… «·’€Ì—…
'            tCard.AddNew
'            tCard!Right = MyMeasure(5.4) - MyMeasure(0.2) + MyMeasure(0.3)
'            tCard!Top = MyMeasure(2.9) - MyMeasure(1.15) - MyMeasure(nDiffer) + MyMeasure(0.3)
'            tCard!Width = MyMeasure(1)
'            tCard!Height = MyMeasure(1.1)
'            tCard!Text = TurnValue(RetPhoto_s(grid1.TextMatrix(I, 7)), "", Null)
'            tCard!isPhoto = True
'            tCard!CardNo = nCard
'            tCard.Update
'        End If
'        '«· ÊﬁÌ⁄
        tCard.AddNew
        tCard!Right = MyMeasure(6.5) - MyMeasure(0.3) + MyMeasure(0.3) - MyMeasure(0.3)
        tCard!Top = MyMeasure(3.3) - MyMeasure(0.2) - MyMeasure(nDiffer) + MyMeasure(0.1) + MyMeasure(0.3)
        tCard!Width = MyMeasure(1.9)
        tCard!Height = MyMeasure(0.9)
        tCard!Text = TurnValue(App.Path & "\sign2.gif", "", Null)
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        
        For i2 = 1 To 10
            '«”„ —∆Ì” ‰«œÌ «·«”ﬂ‰œ—Ì…
            tCard.AddNew
            tCard!Right = MyMeasure(6.7) - MyMeasure(0.8) - MyMeasure(0.3)
            tCard!Top = MyMeasure(4.5) - MyMeasure(0.5) - MyMeasure(nDiffer) + MyMeasure(0.3)
            tCard!Width = MyMeasure(3.1)
            tCard!Height = 1000
            tCard!FontName = "Arial"
            tCard!FontBold = True
            tCard!ForeColor = &H800000
            tCard!FontSize = 9
            tCard!TextAlign = taCenterTop
            tCard!Text = "√.œ.„/Â‘«„ ”⁄ÊœÌ"
            tCard!CardNo = nCard
            tCard.Update
        Next
    End If
Next
prog1.Visible = False
tCard.Requery
doprint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Sub myProc()
ActiveControl.Text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
Unload oSearch
End Sub
Private Sub checkPhoto()
Dim aPrint As Variant
With grid1
prog1.Value = 0
prog1.Visible = True
For I = 1 To grid1.rows - 1
    prog1.Value = Round(I / (grid1.rows - 1), 2) * 100
    If Not validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) Then grid1.Cell(flexcpForeColor, I, 0, I, .Cols - 1) = vbRed
    aPrint = Printed(.TextMatrix(I, 0))
    grid1.TextMatrix(I, 6) = Format(retFlag(aPrint, "date"), "yyyy/mm/dd")
    If IsDate(grid1.TextMatrix(I, 6)) Then .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &HE0E0E0
Next
prog1.Visible = False
End With
End Sub
Private Sub xDown_Change()
' addSetting "down", Val(xDown.Text), cFilePrint
End Sub
Private Sub xSeason_Validate(Cancel As Boolean)
'CmdPrint.Enabled = (Val(xSeason.Text) = 2010)
End Sub
Private Sub myloadGrd()
Dim loctable As New ADODB.Recordset, aDamage As Variant, aDamageOnly As Variant, aMember As Variant
loctable.Open "select * from FILE3_20 WHERE ISCARD = 1 OR ISDAMAGE = 1", con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    aDamage = AddFlag(aDamage, loctable!CODE, loctable!Desca)
    If loctable!isDamage Then
        aDamageOnly = AddFlag(aDamageOnly, "CODE", loctable!CODE)
        aDamageOnly = AddFlag(aDamageOnly, "DESCA", loctable!Desca)
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing


Dim GRDTABLE As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, I As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT FILE3_10.*,DEGREE_GROUP_CODES.DESCA AS DEGREE_GROUP_DESCA, CLASS_CODES.DESCA AS CLASS_DESCA,DEGREE_CODES.DESCA AS DEGREE_DESCA,DEGREE_CODES.DATE AS DATE_END,DEGREE_CODES.[GROUP] AS GROUP_CODE FROM FILE3_10 LEFT JOIN CLASS_CODES ON FILE3_10.CLASS = CLASS_CODES.CODE INNER JOIN DEGREE_CODES ON FILE3_10.DEGREE = DEGREE_CODES.CODE INNER JOIN FILE6_40H ON FILE3_10.CODE = FILE6_40H.CODE INNER JOIN DEGREE_GROUP_CODES ON DEGREE_CODES.[GROUP] = DEGREE_GROUP_CODES.CODE"
If ValidInt(xCode1.Text) Then
    cString = cString & turn(cString) & " FILE3_10.CODE  " & IIf(IsNumeric(xCode2.Text), " >= ", " = ") & xCode1.Text
End If

If IsDate(xDate1.Text) Then
    cString = cString & turn(cString) & "FILE6_40H.DATE >= " & DateSq(xDate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & "FILE6_40H.DATE <= " & DateSq(xDate2.Text)
End If

If ValidInt(xCode2.Text) Then
    cString = cString & turn(cString) & " FILE3_10.CODE <= " & xCode2.Text
End If

cString = cString & " ORDER BY FILE3_10.DESCA"

Set GRDTABLE = New ADODB.Recordset
With grid1
GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (GRDTABLE.EOF And GRDTABLE.BOF) Then
    GRDTABLE.MoveLast
    nRecordcount = GRDTABLE.RecordCount
    GRDTABLE.MoveFirst
End If
prog1.Visible = True
prog1.Value = 0
Dim nFound As Long
Do Until GRDTABLE.EOF
    I = I + 1
    bAddRow = .FindRow(Val(GRDTABLE!CODE & ""), , 0) = -1
    aPrint = Printed(GRDTABLE!CODE)
    aPaid = LastDoc_card(GRDTABLE!CODE)
    
    If bAddRow And xPrinted.Value = 1 Then
        bAddRow = IsEmpty(aPrint)
        If (Not IsEmpty(aPaid)) And (Not IsEmpty(aPrint)) And xDamage.Value = 1 Then
            If retFlag(aDamage, retFlag(aPaid, "ITEM")) <> "" And (Not retFlag(aPaid, "closed")) Then
                bAddRow = True
            End If
        End If
    End If
    
    If bAddRow And xPaid.Value = 1 Then
        bAddRow = Not IsEmpty(aPaid)
    End If
        
    If bAddRow Then
        prog1.Value = Round(I / (nRecordcount), 2) * 100
               
        .AddItem ""
        .TextMatrix(grid1.rows - 1, 0) = GRDTABLE!CODE
        .TextMatrix(grid1.rows - 1, 1) = GRDTABLE!Desca
        .TextMatrix(grid1.rows - 1, 2) = GRDTABLE!DEGREE_desca
        .TextMatrix(grid1.rows - 1, 3) = GRDTABLE!CLASS_DESCA & ""
        .TextMatrix(grid1.rows - 1, 4) = myFormat(GRDTABLE!date_end)
        
        .TextMatrix(grid1.rows - 1, 5) = Format(retFlag(aPaid, "form_no"))
        .TextMatrix(grid1.rows - 1, 6) = Format(retFlag(aPrint, "date"), "YYYY/M/D")
        If retFlag(aDamage, retFlag(aPaid, "ITEM")) <> "" And xPrinted.Value = 1 Then
            .TextMatrix(grid1.rows - 1, 7) = retFlag(aPaid, "ITEM")
            .TextMatrix(grid1.rows - 1, 8) = retFlag(aDamage, retFlag(aPaid, "ITEM"))
        ElseIf xPrinted.Value = 0 And (Not IsEmpty(aPrint)) Then
            .TextMatrix(grid1.rows - 1, 7) = retFlag(aDamageOnly, "CODE")
            .TextMatrix(grid1.rows - 1, 8) = retFlag(aDamageOnly, "DESCA")
        End If
        .TextMatrix(grid1.rows - 1, 9) = GRDTABLE!GROUP_CODE & ""
        .TextMatrix(grid1.rows - 1, 10) = GRDTABLE!DEGREE_GROUP_DESCA & ""
        .TextMatrix(grid1.rows - 1, .Cols - 1) = 0
    End If
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing

prog1.Visible = False
Me.MousePointer = 0
If grid1.rows > 1 Then
    grid1.Select 1, 0, 1, 1
    grid1.Sort = flexSortGenericAscending
End If
CalcTotals
End With
End Sub
Private Sub Fixgrd()
With grid1
    
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄÷Ê"
    .TextMatrix(0, 1) = "«·«”„"
    .TextMatrix(0, 2) = "«·›—ﬁ…"
    .TextMatrix(0, 3) = "«·‘⁄»…"
    .TextMatrix(0, 4) = "Ì‰ ÂÏ ›Ï"
    
    .TextMatrix(0, 5) = "—ﬁ„ «·«Ì’«·"
    .TextMatrix(0, 6) = " «—ÌŒ «·ÿ»«⁄…"
    .TextMatrix(0, 7) = "‰Ê⁄ «·ﬂ«—‰ÌÂ"
    .TextMatrix(0, 8) = "‰Ê⁄ «·ﬂ«—‰ÌÂ"
    .TextMatrix(0, 9) = "«·„Ã„Ê⁄…"
    .TextMatrix(0, 10) = "«·„Ã„Ê⁄…"
    .TextMatrix(0, 11) = "«Œ Ì«—"
            
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1400
    .ColWidth(5) = 1000
    .ColWidth(6) = 1400
    .ColWidth(7) = 1000
    .ColWidth(8) = 1600
    .ColWidth(10) = 1900
    .ColHidden(9) = True
    .ColHidden(7) = True
    
    For I = 0 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    .ColDataType(0) = flexDTLong
    .ColDataType(.Cols - 1) = flexDTBoolean
End With
End Sub
Private Sub FixClass()

Dim cString As String, cSave As String
cString = "Select * from class_codes"
If xDegree.MatchedWithList Then cString = cString & turn(cString) & "[degree] = " & xDegree.BoundText
If xType.MatchedWithList Then cString = cString & turn(cString) & "[type] = " & xType.BoundText
If xLang.MatchedWithList Then cString = cString & turn(cString) & "[lang] = " & xLang.BoundText
If xReg.MatchedWithList Then cString = cString & turn(cString) & "[reg] = " & xReg.BoundText
cString = cString & cOrderBy
If Trim(UCase(data1.Recordset.Source)) <> Trim(UCase(cString)) Then
    cSave = xClass.BoundText
    Set data1.Recordset = myRecordSet(cString, con)
    xClass.BoundText = cSave
    If Not xClass.MatchedWithList Then xClass.BoundText = ""
End If
End Sub

Private Sub xType_Click(Area As Integer)
FixClass
End Sub
Private Sub CalcTotals()
Dim nAll As Long, nPhoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(3).Text = ""
StatusBar1.Panels(2).Text = ""
StatusBar1.Panels(1).Text = ""
If grid1.rows = 1 Then Exit Sub
For I = 0 To grid1.rows - 1
    nAll = nAll + 1
    If validPhoto(RetPhoto_s(grid1.TextMatrix(I, 0))) Then nPhoto = nPhoto + 1
Next
nPhoto2 = ((Val(xRow.Text) - 1) * 2) + (Val(xCol.Text) - 1)
nPages = Fix(nPhoto / 10)
If nPhoto > 10 Then nLeft = nPhoto Mod 10
StatusBar1.Panels(3).Text = "⁄œœ «·”Ã·«  : " & nAll
StatusBar1.Panels(2).Text = "⁄œœ «·”Ã·«  »’Ê— : " & nPhoto
StatusBar1.Panels(1).Text = "⁄œœ «·’›Õ«  : " & nPages
If nrest > 0 Then StatusBar1.Panels(3).Text = StatusBar1.Panels(3).Text & turn(StatusBar1.Panels(3).Text, " ") & nrest & " ’Ê—…"
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub
Private Sub xSeason_GotFocus()
myGotFocus xSeason
End Sub
Private Sub xSeason_LostFocus()
myLostFocus xSeason
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xAppend_GotFocus()
myGotFocus xAppend
End Sub
Private Sub xAppend_LostFocus()
myLostFocus xAppend
End Sub
Private Sub xCode2_GotFocus()
myGotFocus xCode2
End Sub
Private Sub xCode2_LostFocus()
myLostFocus xCode2
End Sub
Private Sub xCode1_GotFocus()
myGotFocus xCode1
End Sub
Private Sub xCode1_LostFocus()
myLostFocus xCode1
xcode_desca.Caption = ""
If Not ValidInt(xCode1.Text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from FILE3_10 where code = " & xCode1.Text)
If Not IsEmpty(aRet) Then
    xcode_desca.Caption = retFlag(aRet, "DESCA") & ""
End If
End Sub
Private Sub xDown_GotFocus()
myGotFocus xDown
End Sub
Private Sub xDown_LostFocus()
myLostFocus xDown
End Sub
Private Sub xRight_GotFocus()
myGotFocus xRight
End Sub
Private Sub xRight_LostFocus()
myLostFocus xRight
End Sub
Private Sub xCol_GotFocus()
myGotFocus xCol
End Sub
Private Sub xCol_LostFocus()
myLostFocus xCol
End Sub
Private Sub xRow_GotFocus()
myGotFocus xRow
End Sub
Private Sub xRow_LostFocus()
myLostFocus xRow
End Sub
Private Function Printed(pCode As String) As Variant
Printed = GetFields("SELECT TOP 1 FILE4_30.* FROM FILE4_30  WHERE FILE4_30.CODE =  " & pCode & " AND FILE4_30.[YEAR] = " & sSeason_Student & " ORDER BY FILE4_30.DATE DESC", con)
End Function
Private Function LastDoc_card(pMember As String) As Variant
cString = "SELECT TOP 1 FILE6_40H.*,FILE6_40.CODE AS ITEM FROM FILE6_40H INNER JOIN FILE6_40 ON FILE6_40H.DOC_NO = FILE6_40.DOC_NO" & _
          " WHERE FILE6_40.MEMBER = " & pMember & " and closed = 0 AND FILE6_40H.SEASON =  " & sSeason_Student & _
           " ORDER BY FILE6_40H.DATE DESC,FILE6_40H.DOC_NO DESC"
LastDoc_card = GetFields(cString, con)
End Function


