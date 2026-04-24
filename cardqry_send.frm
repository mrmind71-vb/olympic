VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cardqry_sendfrm 
   Caption         =   "ÿ»«⁄… «·ﬂ«—‰ÌÂ« "
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18660
   FillColor       =   &H00008000&
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   3015
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   270
      Width           =   1680
   End
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "ÿ»«⁄… «·–Ì ·„ Ìÿ»⁄ ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   585
      Width           =   2265
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«ŸÂ«— «·«⁄÷«¡ »œÊ‰ ’Ê— ›ﬁÿ"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7785
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   585
      Width           =   2985
   End
   Begin VB.Frame Frame10 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   945
      Width           =   10815
      Begin VB.CheckBox chkMobil 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„Õ„Ê· ›ﬁÿ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chkIgPhoto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄œ„ «· Õﬁﬁ „‰ «·’Ê—"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   270
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox ckhPrintMain 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ÿ»«⁄… «·⁄÷Ê «·—∆Ì”Ì ›ﬁÿ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6075
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   270
         Width           =   2445
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ÿ»«⁄… „Œ ’—…"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   270
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«ŸÂ«— «⁄÷«¡ »œÊ‰ ’Ê—"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   270
         Width           =   2220
      End
   End
   Begin VB.Frame Frame9 
      Height          =   735
      Left            =   12555
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   -45
      Width           =   1365
      Begin VB.CommandButton cmdToFlash 
         Height          =   510
         Left            =   45
         Picture         =   "cardqry_send.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   13950
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   -45
      Width           =   4650
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
         Left            =   2685
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry_send.frx":2341
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry_send.frx":47AD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
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
      Height          =   735
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   8955
      Width           =   7665
      Begin ComctlLib.ProgressBar prog1 
         Height          =   510
         Left            =   45
         TabIndex        =   34
         Top             =   180
         Visible         =   0   'False
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   900
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
      Height          =   735
      Left            =   13635
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   8955
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
         TabIndex        =   30
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label6 
         Caption         =   "«·’›"
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
         TabIndex        =   32
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "«·⁄„Êœ"
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
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   31
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
      Height          =   735
      Left            =   15975
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8955
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label8 
         Caption         =   "«”›·"
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
         TabIndex        =   27
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label9 
         Caption         =   "Ì„Ì‰"
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
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8955
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
         TabIndex        =   12
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
         TabIndex        =   11
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
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1620
      Width           =   5865
      Begin VB.CommandButton cmdAdd 
         Height          =   555
         Left            =   4455
         Picture         =   "cardqry_send.frx":7047
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   45
         Picture         =   "cardqry_send.frx":9539
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1410
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   555
         Left            =   3105
         Picture         =   "cardqry_send.frx":BD24
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1320
      End
      Begin Threed.SSCommand cmdprintrep 
         Height          =   555
         Left            =   1485
         TabIndex        =   8
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
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
         Picture         =   "cardqry_send.frx":E14E
         Caption         =   "ÿ»«⁄…  ﬁ—Ì—"
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
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
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
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
      Height          =   330
      Left            =   -1125
      Top             =   270
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      BackColor       =   16777215
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
      Left            =   -360
      Top             =   540
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
      Left            =   -585
      Top             =   630
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
      Left            =   -315
      Top             =   315
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   420
      Left            =   -1215
      Top             =   -270
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
      Left            =   -1305
      Top             =   45
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
      Height          =   1770
      Left            =   10980
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   630
      Width           =   7575
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
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1185
      End
      Begin VB.TextBox xAppend 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   945
         Width           =   1635
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   945
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo xCompany 
         Height          =   330
         Left            =   3285
         TabIndex        =   41
         Top             =   1305
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdYearPaid 
         Height          =   735
         Left            =   135
         TabIndex        =   45
         Top             =   945
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1296
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "„”œœ „Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "«·‘—ﬂ…"
         Height          =   240
         Index           =   2
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1350
         Width           =   960
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
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   690
      End
      Begin VB.Label xcode_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   5100
      End
      Begin VB.Label Label3 
         Caption         =   " ⁄÷Ê  «»⁄ "
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
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   630
         Width           =   990
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   19
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
         Index           =   0
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   825
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "cardqry_send.frx":10752
      Height          =   6405
      Left            =   90
      TabIndex        =   37
      Top             =   2430
      Width           =   18510
      _cx             =   32650
      _cy             =   11298
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8910
      Width           =   2760
      Begin VB.CommandButton Command1 
         Caption         =   "«Œ Ì«— «·ﬂ·"
         Height          =   555
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   2670
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "„Ê”„ «·ÿ»«⁄…"
      Height          =   780
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1620
      Width           =   2175
      Begin Threed.SSCommand cmdYear 
         Height          =   420
         Left            =   90
         TabIndex        =   40
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«÷€ÿ ·«Œ Ì«— «·Õ«·Ì…"
         ButtonStyle     =   3
      End
   End
   Begin VB.Frame fmManager 
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1620
      Width           =   2760
      Begin VB.CommandButton cmdDelPrinted 
         Caption         =   "Õ–› «·„ÿ»Ê⁄"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1305
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1410
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
         Height          =   555
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cardqry_send.frx":10765
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grdExcel 
      Bindings        =   "cardqry_send.frx":12AC8
      Height          =   6405
      Left            =   -90
      TabIndex        =   47
      Top             =   3510
      Visible         =   0   'False
      Width           =   18510
      _cx             =   32650
      _cy             =   11298
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
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
End
Attribute VB_Name = "cardqry_sendfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFileSave As String, cFileSave2 As String, cFilePrint As String
Dim oSearch As New Search, oSearchYear As New Search, oSearchYearPaid As New Search
Dim con As New ADODB.Connection
Dim printTable As New ADODB.Recordset
Private Sub cmdAdd_Click()
'checkErr
myLoadGrd
StatusBar1.Panels(1).text = grid1.rows - 1
If Check2.Value = 0 Then
    On Error Resume Next
    grid1.SaveGrid cFileSave, flexFileData
    Err.Clear
    cmdPrint.Enabled = (grid1.rows > 1)
    checkPhoto
End If
End Sub

Private Sub cmdDelPrinted_Click()
If MsgBox("Õ–› «·ÿ»«⁄… ··ﬂ· øø", vbDefaultButton2 + vbOKCancel) <> vbOK Then
    Exit Sub
End If
nDeleted = DeletePrint
If nDeleted > 0 Then
    MsgBox " „ Õ–› " & nDeleted & "„‰ «·ÿ»«⁄…"
    checkPhoto
End If
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
sCaption = Me.Caption

For i = 0 To grid1.Cols - 1
    If grid1.ColHidden(i) Then grid1.ColHidden(i) = False
Next

If chkIgPhoto.Value = 0 Then
    For i = 1 To grid1.rows - 1
        Me.Caption = sCaption & "”Ã· " & i & " „‰ " & grid1.rows - 1
        If Check1.Value = 1 Then
            If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
                grid1.RowHidden(i) = True
            End If
        Else
            If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
                grid1.RowHidden(i) = True
            End If
        End If
    Next
End If

If chkMobil.Value = 1 Then
    grid1.ColHidden(0) = True
    grid1.ColHidden(1) = True
    grid1.ColHidden(1) = True
    grid1.ColHidden(3) = True
    grid1.ColHidden(5) = True
    grid1.ColHidden(6) = True
    grid1.ColHidden(7) = True
    grid1.ColHidden(8) = True
Else
    grid1.ColHidden(9) = True
End If

ToFileExel2 grid1, Array(2), , , , 1, , , , , , Me

For i = 1 To grid1.rows - 1
    If grid1.RowHidden(i) Then grid1.RowHidden(i) = False
Next
For i = 0 To grid1.Cols - 1
    grid1.ColHidden(i) = False
Next

End Sub

Private Sub cmdGo_Click()

End Sub

Private Sub cmdPrint_Click()
If grid1.rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»⁄Â«"
    Exit Sub
End If

If Val(xRow.text) > 5 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.text) > 2 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If

If retFlag(aSet, "branch", 1) = 1 Then
    If Not doPrint Then
        MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
        Exit Sub
    End If
Else
    If Not doprint2 Then
        MsgBox "·«  ÊÃœ ”Ã·«  ··ÿ»«⁄…"
        Exit Sub
    End If
End If

Set CardPrintNew.myForm = Me
'CardPrintNew.pDevice = RetSetting("printer2", tempPath & turn(tempPath, "\") & "printers.txt")
'CardPrintNew.pLand = True
CardPrintNew.PrintArray
CardPrintNew.Show 1

grid1.SaveGrid cFileSave2, flexFileData
If MsgBox(" „  «·ÿ»«⁄… ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    SavePrinted
    For i = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 Then
            grid1.RemoveItem i
        End If
    Next
    checkPhoto
End If
'grid1.SaveGrid cFilesave, flexFileData
End Sub
Private Sub CmdExit_Click()
Unload Me
Set cardqryfrm = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdClear_Click()
grid1.rows = 1
End Sub
Private Sub cmdMember_Click()
member.Show 1
End Sub
Private Sub CmdLastFillGrd_Click()
Dim fs As New FileSystemObject
If fs.FileExists(cFileSave) Then
    grid1.LoadGrid cFileSave, flexFileData
    If grid1.rows > 1 Then cmdPrint.Enabled = True
    checkPhoto
End If
End Sub

Private Sub cmdSave_Click()
If MsgBox(" „  «·ÿ»«⁄… ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    SavePrinted
    For i = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 Then
            grid1.RemoveItem i
        End If
    Next
    checkPhoto
End If
End Sub
Private Sub cmdToFlash_Click()
Dim fs As New FileSystemObject
Dim myPathPhoto As String
myPathPhoto = App.Path & "\Flash\Photos"
MyCreateFolder myPathPhoto
With grdExcel
grdExcel.rows = 1

prog1.Visible = True
prog1.Value = 0

'.TextMatrix(0, 0) = "#"
'.TextMatrix(0, 1) = "«··ﬁ»"
'
'.TextMatrix(0, 2) = "«·≈”„"
'.TextMatrix(0, 3) = "‰Ê⁄ «·⁄÷ÊÌ…"
'
'.TextMatrix(0, 4) = "—ﬁ„ «·⁄÷ÊÌ…"
'.TextMatrix(0, 5) = "«·’Ê—…"
'
'    .TextMatrix(0, 0) = "—ﬁ„ «·⁄÷Ê"
'    .TextMatrix(0, 1) = "—ﬁ„ «· «»⁄"
'    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
'    .TextMatrix(0, 3) = "«··ﬁ»"
'    .TextMatrix(0, 4) = "«·«”„"
'    .TextMatrix(0, 5) = "‰Ê⁄ «·⁄÷ÊÌ…"
'    .TextMatrix(0, 6) = "—ﬁ„ «·„” ‰œ"
'    .TextMatrix(0, 7) = " «—ÌŒ «·„” ‰œ"
'    .TextMatrix(0, 8) = "‰Ê⁄ «·„” ‰œ"
'    .TextMatrix(0, 9) = "—ﬁ„ «·„Õ„Ê·"

On Error GoTo myerror
For i = 1 To grid1.rows - 1
    prog1.Value = mRound(i / (grid1.rows - 1), 2) * 100
    Me.Caption = "”Ã· " & i & " „‰ " & grid1.rows - 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
        .AddItem ""
        .TextMatrix(.rows - 1, 0) = .rows - 1
        .TextMatrix(.rows - 1, 1) = grid1.TextMatrix(i, 3)
        .TextMatrix(.rows - 1, 2) = grid1.TextMatrix(i, 4)
        .TextMatrix(.rows - 1, 3) = grid1.TextMatrix(i, 5)
        .TextMatrix(.rows - 1, 4) = grid1.TextMatrix(i, 2)
        .TextMatrix(.rows - 1, 5) = grid1.TextMatrix(i, 2)
        If Check1.Value = 0 Then
            fs.CopyFile retPhoto(grid1.TextMatrix(i, 2)), myPathPhoto & "\" & grid1.TextMatrix(i, 2) & ".Jpg"
        End If
    End If
Next
prog1.Visible = False
prog1.Value = 0
Dim sFile As String
sFile = App.Path & "\flash\data.xlsx"
ToFileExel2 grdExcel, , , , , 1, , , , 12, , Me
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdYear_Click()
Set oSearchYear = New Search
Years_LookupAll Me, oSearchYear
End Sub
Private Sub cmdYearPaid_Click()
Set oSearchYearPaid = New Search
Years_LookupAll Me, oSearchYearPaid, , cmdYearPaid.Tag <> ""
'cmdAdd_Click
End Sub
Private Sub Command1_Click()
For i = 1 To grid1.rows - 1
    grid1.TextMatrix(i, grid1.Cols - 1) = 1
Next
End Sub

Private Sub cmdprintrep_Click()
Set PrintGrdNew.myForm = Me
Dim i As Long
For i = 1 To grid1.rows - 1
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 0))) Then
        grid1.RowHidden(i) = True
    End If
Next
grid1.ColHidden(grid1.Cols - 1) = True
PrintGrdNew.doPrint grid1, 0.8, -3, "ÿ»«⁄… «·«⁄÷«¡", , , , False, False, 9
PrintGrdNew.Show 1
grid1.ColHidden(grid1.Cols - 1) = False
End Sub

Private Sub Command2_Click()
For i = 1 To grid1.rows - 1
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.RowHidden(i) = False
        grid1.Cell(flexcpBackColor, i, 0, i, grid1.Cols - 1) = vbRed
    End If
Next
End Sub

Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    grid1.LoadGrid cFileSave2, flexFileData
    For i = grid1.rows - 1 To 1 Step -1
        If Val(grid1.TextMatrix(i, grid1.Cols - 1)) = 0 Then
            grid1.RemoveItem i
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
FixgrdExcel

Set DATA1.Recordset = myRecordSet("SELECT CODE,DESCA FROM COMPANY_CODES ORDER BY CODE", con)
Set xCompany.RowSource = DATA1
xCompany.ListField = "Desca"
xCompany.BoundColumn = "Code"

'LoadText Me

cmdYear.Caption = ArbString(retFlag(aSeasonPrint, "desca"))
cmdYear.Tag = retFlag(aSeasonPrint, "code")
'xPrinted.Value = 1
fmManager.Visible = bopt1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveText Me
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'grid1.SaveGrid cFileSave, flexFileData
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
    checkPhoto
End If
End With
End Sub
Private Sub xCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    xUnCode.Caption = ""
    If xCode.text = "" Then Exit Sub
    xUnCode.ForeColor = -2147483630
    If Val(unMyCodeBar(xCode.text, "1")) <> 1 Then
        xUnCode.Caption = "Error"
        xUnCode.ForeColor = vbRed
    Else
        xUnCode.Caption = unMyCodeBar(xCode.text)
    End If
    myGotFocus xCode
End If
End Sub

Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
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
    MemberLookupAll Me, oSearch
End If
End Sub
Private Sub xCode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub
Private Sub xCODE2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    MemberLookupAll Me, oSearch
End If
End Sub
Private Function CountGrid() As Integer
With grid1
For i = 1 To grid1.rows - 1
    CountGrid = CountGrid + 1
Next
End With
End Function
Private Sub countPrint()
nCountPrint = 0
With grid1
For i = 1 To .rows - 1
   If .TextMatrix(i, .Cols - 1) = True Then nCountPrint = nCountPrint + 1
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
Dim aInsert As Variant, i As Long
con.BeginTrans
prog1.Value = 0
prog1.Visible = True
For i = 1 To .rows - 1
    prog1.Value = Round(i / (grid1.rows - 1), 2) * 100
   If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 And Val(grid1.TextMatrix(i, grid1.Cols - 2)) = 0 Then
        aInsert = AddFlag(Empty, "MEMBER", grid1.TextMatrix(i, 0))
        aInsert = AddFlag(aInsert, "CODE", addvalue(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "DESCA2", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "TITLE", addstring(grid1.TextMatrix(i, 5)))
        
        aInsert = AddFlag(aInsert, "RELATION", addstring(grid1.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "RELATION_DESCA", addstring(grid1.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "TYPE_DESCA", addstring(grid1.TextMatrix(i, 8)))
        
        aInsert = AddFlag(aInsert, "DATE", addDate(Now))
        aInsert = AddFlag(aInsert, "YEAR", addvalue(cmdYear.Tag))
        
        con.Execute addInsert(aInsert, "file4_10")
        
        aInsert = AddFlag(Empty, "DATE_PRINT", addDate(Format(Date, "YYYY-MM-DD HH:NN")))
        If grid1.TextMatrix(i, 1) = "" Then
            con.Execute addUpdate(aInsert, "FILE1_10", "FILE1_10.CODE = " & grid1.TextMatrix(i, 0))
        Else
            con.Execute addUpdate(aInsert, "FILE1_11", "FILE1_11.MEMBER = " & grid1.TextMatrix(i, 0) & " AND FILE1_11.CODE = " & grid1.TextMatrix(i, 1))
        End If
    End If
Next
con.CommitTrans
End With
prog1.Value = 0
prog1.Visible = False
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
prog1.Value = 0
prog1.Visible = False
End Sub
Function eofGrd(cId) As Boolean
eofGrd = (grid1.FindRow(cId, , 0) = -1)
End Function
Private Function doPrint() As Boolean
Dim nDiffer As Double
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cRightMargin) = MyMeasure(-0.1)
SettingArray(cCardWidth) = MyMeasure(0)
SettingArray(cCardHeight) = MyMeasure(5.755)
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(8.6)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
nRows = SettingArray(cRows)
nspace = 0.62
nup = 0.3

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > nRows, 1, nRow)
    blastrow = (nRow = nRows)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡


prog1.Value = 0
prog1.Visible = True

For i = 1 To grid1.rows - 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 And Val(grid1.TextMatrix(i, grid1.Cols - 2)) = 0 Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 0.75
        
        nTop = MyMeasure(2.1)
        nVSpace = MyMeasure(0.65)
        
        nCard = nCard + 1
        
        blastrow = (nRow = nRows)
        nDiffer = 0.75
        
        nTop = MyMeasure(2.1)
        nVSpace = MyMeasure(0.65)
        
        tCard.AddNew
        tCard!Right = MyMeasure(1)
        tCard!Top = nTop - MyMeasure(0.95)
        tCard!Width = MyMeasure(5)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = TurnValue(IIf(grid1.TextMatrix(i, 8) = "⁄÷Ê ⁄«„·", &HC00000, &HC00000), "", Null)
        tCard!FontSize = 13
        tCard!text = grid1.TextMatrix(i, 8)
        tCard!CardNo = nCard
        tCard!TextAlign = taCenterTop
        tCard.Update
        
        '«··ﬁ»
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 1) - MyMeasure(0.55) - MyMeasure(0.3)
        tCard!Width = MyMeasure(5)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 10
        tCard!text = TurnValue(grid1.TextMatrix(i, 5))
        tCard!CardNo = nCard
        tCard.Update
        
        'Next
        ' «·«”„
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 2) - MyMeasure(0.45) - MyMeasure(0.6)
        tCard!Width = MyMeasure(5.8)
        tCard!Height = MyMeasure(0)
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = &H800000
        tCard!FontSize = 10
        tCard!text = TurnValue(IIf(grid1.TextMatrix(i, 6) = "", grid1.TextMatrix(i, 3), grid1.TextMatrix(i, 4)))
        tCard!CardNo = nCard
        tCard.Update
        
        ' ﬂ·„… —ﬁ„ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 3) - MyMeasure(0.9)
        tCard!Width = MyMeasure(2.3)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 10
        tCard!text = ArbString("—ﬁ„ «·⁄÷ÊÌ… :")
        tCard!CardNo = nCard
        tCard.Update
        
         tCard.AddNew
         tCard!Right = MyMeasure(3.05)
         tCard!Top = nTop + (nVSpace * 3) - MyMeasure(0.9)
         tCard!Width = 0
         tCard!Height = 0
         tCard!TextAlign = taLeftTop
         tCard!FontName = "GE SS two Medium"
         tCard!FontBold = True
         tCard!ForeColor = &HFF&
         tCard!FontSize = 11
         tCard!text = TurnValue(ArbString(myRev(grid1.TextMatrix(i, 2))))
         tCard!CardNo = nCard
         tCard.Update
    
         ' Ì‰ ÂÌ ›Ì
         tCard.AddNew
         tCard!Right = MyMeasure(1.2)
         tCard!Top = nTop + (nVSpace * 4) - MyMeasure(0.7)
         tCard!Width = 0
         tCard!Height = 0
         tCard!FontName = "GE SS two Medium"
         tCard!FontBold = True
         tCard!ForeColor = &HFF&
         tCard!FontSize = 9
         'tCard!Text = "Ì‰ ÂÌ ›Ì " & IIf(RetMember(Grid1.TextMatrix(I, 6), "DEBIT"), Format("31/12/2014", "yyyy/m/d"), Format("31/12/2019", "yyyy/m/d"))
         tCard!text = "Ì‰ ÂÌ ›Ì " & Format("30/6/2023", "yyyy/m/d")
         tCard!CardNo = nCard
         tCard.Update
        
        
        'ﬂ·„… „œÌ— ⁄«„ «·‰«œÌ
        tCard.AddNew
        tCard!Right = MyMeasure(6) - MyMeasure(0.4)
        tCard!Top = MyMeasure(5.2) - MyMeasure(0.8) - MyMeasure(0.4) - MyMeasure(0.1) + MyMeasure(0.05)
        tCard!Width = MyMeasure(3)
        tCard!Height = 1000
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &HFF&
        tCard!FontSize = 8
        tCard!TextAlign = &H800000
        tCard!TextAlign = taCenterTop
        tCard!text = "—∆Ì” „Ã·” «·«œ«—…"
        tCard!CardNo = nCard
        tCard.Update
            
    '«”„ „œÌ— ⁄«„ «·‰«œÌ
        tCard.AddNew
        tCard!Right = MyMeasure(6) - MyMeasure(0.3)
        tCard!Top = MyMeasure(4.8) - MyMeasure(0.4) - MyMeasure(0.1)
        tCard!Width = MyMeasure(3)
        tCard!Height = MyMeasure(1)
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &H800000
        tCard!FontSize = 8
        tCard!TextAlign = taCenterTop
        tCard!text = "„Õ„œ „’Ì·ÕÌ"
        tCard!CardNo = nCard
        tCard.Update
        
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.26) - MyMeasure(0.1)
        tCard!Top = MyMeasure(2.14) - MyMeasure(0.25) - MyMeasure(0.41) + MyMeasure(0.17)
        tCard!Width = MyMeasure(2.15) * 0.85
        tCard!Height = MyMeasure(2.65) * 0.85
        tCard!text = retPhoto(grid1.TextMatrix(i, 2))
        'tCard!Text = App.Path & "\" & "logo.jpg"
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        If mRound(grid1.TextMatrix(i, 6)) > 1 Then
            If validPhoto(retPhoto(grid1.TextMatrix(i, 0))) Then
                '«·’Ê—… «·’€Ì—…
                tCard.AddNew
                tCard!Right = MyMeasure(5.1) + MyMeasure(0.25)
                tCard!Top = MyMeasure(2) + MyMeasure(3) - MyMeasure(1.2) - MyMeasure(0.6) - MyMeasure(0.4) - MyMeasure(0.15) + MyMeasure(0.12)
                tCard!Width = MyMeasure(1.05) * 0.95
                tCard!Height = MyMeasure(1.2) * 0.95
                tCard!text = TurnValue(retPhoto(grid1.TextMatrix(i, 0)), "", Null)
                tCard!isPhoto = True
                tCard!CardNo = nCard
                tCard.Update
            End If
        End If
    End If
Next
prog1.Visible = False
tCard.Requery
doPrint = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Private Function doprint2() As Boolean
Dim nDiffer As Double
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cRightMargin) = MyMeasure(-0.1)
SettingArray(cCardWidth) = MyMeasure(0)
SettingArray(cCardHeight) = MyMeasure(5.755)
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(8.6)

contemp.Execute "delete * From Card"

Dim tCard As New ADODB.Recordset
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
nRows = SettingArray(cRows)
nspace = 0.62
nup = 0.3

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > nRows, 1, nRow)
    blastrow = (nRow = nRows)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡


prog1.Value = 0
prog1.Visible = True

For i = 1 To grid1.rows - 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) And Val(grid1.TextMatrix(i, grid1.Cols - 1)) <> 0 And Val(grid1.TextMatrix(i, grid1.Cols - 2)) = 0 Then
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > nRows, 1, nRow)
        blastrow = (nRow = nRows)
        nDiffer = 0.75
        
        nTop = MyMeasure(2.1)
        nVSpace = MyMeasure(0.65)
        
        nCard = nCard + 1
        
        blastrow = (nRow = nRows)
        nDiffer = 0.75
        
        nTop = MyMeasure(2.1)
        nVSpace = MyMeasure(0.65)
        
        
        tCard.AddNew
        tCard!Right = MyMeasure(0.45)
        tCard!Top = nTop - MyMeasure(0.95) - MyMeasure(0.8) + MyMeasure(0.2) + MyMeasure(0.2)
        tCard!Width = MyMeasure(7)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &H8000&
        tCard!FontSize = 10
        tCard!text = "‰«œÌ «·« Õ«œ «·”ﬂ‰œ—Ì"
        tCard!CardNo = nCard
        tCard!TextAlign = taRightTop
        tCard.Update
        
        tCard.AddNew
        tCard!Right = MyMeasure(0.45) + MyMeasure(0.4)
        tCard!Top = nTop - MyMeasure(0.95) - MyMeasure(0.3) + MyMeasure(0.15) + MyMeasure(0.2)
        tCard!Width = MyMeasure(7)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &H8000&
        tCard!FontSize = 10
        tCard!text = "›‹‹—⁄ ”„ÊÕ‹‹…"
        tCard!CardNo = nCard
        tCard!TextAlign = taRightTop
        tCard.Update
        
        
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop - MyMeasure(0.95) + MyMeasure(0.7)
        tCard!Width = MyMeasure(7)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = TurnValue(IIf(grid1.TextMatrix(i, 8) = "⁄÷Ê ⁄«„·", &HC00000, &HC00000), "", Null)
        tCard!FontSize = 13
        tCard!text = grid1.TextMatrix(i, 8)
        tCard!CardNo = nCard
        tCard!TextAlign = taCenterTop
        tCard.Update
        
        '«··ﬁ»
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 1) - MyMeasure(0.55) - MyMeasure(0.3) + MyMeasure(0.7)
        tCard!Width = MyMeasure(5)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 10
        tCard!text = TurnValue(grid1.TextMatrix(i, 5))
        tCard!CardNo = nCard
        tCard.Update
        
        'Next
        ' «·«”„
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 2) - MyMeasure(0.45) - MyMeasure(0.6) + MyMeasure(0.7)
        tCard!Width = MyMeasure(5.8)
        tCard!Height = MyMeasure(0)
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = True
        tCard!ForeColor = &H800000
        tCard!FontSize = 10
        tCard!text = TurnValue(IIf(grid1.TextMatrix(i, 6) = "", grid1.TextMatrix(i, 3), grid1.TextMatrix(i, 4)))
        tCard!CardNo = nCard
        tCard.Update
        
        ' ﬂ·„… —ﬁ„ «·⁄÷ÊÌ…
        tCard.AddNew
        tCard!Right = MyMeasure(0.7)
        tCard!Top = nTop + (nVSpace * 3) - MyMeasure(0.9) + MyMeasure(0.7)
        tCard!Width = MyMeasure(2.3)
        tCard!Height = 0
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = vbBlack
        tCard!FontSize = 10
        tCard!text = ArbString("—ﬁ„ «·⁄÷ÊÌ… :")
        tCard!CardNo = nCard
        tCard.Update
        
         tCard.AddNew
         tCard!Right = MyMeasure(3.05)
         tCard!Top = nTop + (nVSpace * 3) - MyMeasure(0.9) + MyMeasure(0.7)
         tCard!Width = 0
         tCard!Height = 0
         tCard!TextAlign = taLeftTop
         tCard!FontName = "GE SS two Medium"
         tCard!FontBold = True
         tCard!ForeColor = &HFF&
         tCard!FontSize = 11
         tCard!text = TurnValue(ArbString(myRev(grid1.TextMatrix(i, 2))))
         tCard!CardNo = nCard
         tCard.Update
    
         ' Ì‰ ÂÌ ›Ì
         tCard.AddNew
         tCard!Right = MyMeasure(1.2)
         tCard!Top = nTop + (nVSpace * 4) - MyMeasure(0.7) + MyMeasure(0.7)
         tCard!Width = 0
         tCard!Height = 0
         tCard!FontName = "GE SS two Medium"
         tCard!FontBold = True
         tCard!ForeColor = &HFF&
         tCard!FontSize = 9
         tCard!text = "Ì‰ ÂÌ ›Ì " & Format("30/6/2023", "yyyy/m/d")
         tCard!CardNo = nCard
         tCard.Update
                
        ' ﬂ·„… „œÌ— ⁄«„ «·‰«œÌ
        tCard.AddNew
        tCard!Right = MyMeasure(6) - MyMeasure(0.25)
        tCard!Top = MyMeasure(5.2) - MyMeasure(0.8) - MyMeasure(0.4) - MyMeasure(0.1) + MyMeasure(0.15) + MyMeasure(0.2)
        tCard!Width = MyMeasure(3)
        tCard!Height = 1000
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &HFF&
        tCard!FontSize = 8
        tCard!TextAlign = &H800000
        tCard!TextAlign = taCenterTop
        tCard!text = "—∆Ì” „Ã·” «·«œ«—…"
        tCard!CardNo = nCard
        tCard.Update
            
        ' «”„ „œÌ— ⁄«„ «·‰«œÌ
        tCard.AddNew
        tCard!Right = MyMeasure(6) - MyMeasure(0.3)
        tCard!Top = MyMeasure(4.8) - MyMeasure(0.4) - MyMeasure(0.1) + MyMeasure(0.15) + MyMeasure(0.2)
        tCard!Width = MyMeasure(3)
        tCard!Height = MyMeasure(1)
        tCard!FontName = "GE SS two Medium"
        tCard!FontBold = False
        tCard!ForeColor = &H800000
        tCard!FontSize = 8
        tCard!TextAlign = taCenterTop
        tCard!text = "„Õ„œ „’Ì·ÕÌ"
        tCard!CardNo = nCard
        tCard.Update
        
        '«·’Ê—… «·ﬂ»Ì—…
        tCard.AddNew
        tCard!Right = MyMeasure(6.26) - MyMeasure(0.1) + MyMeasure(0.2)
        tCard!Top = MyMeasure(2.14) - MyMeasure(0.25) - MyMeasure(0.41) + MyMeasure(0.1) + MyMeasure(0.2)
        tCard!Width = MyMeasure(2.15) * 0.9
        tCard!Height = MyMeasure(2.65) * 0.9
        tCard!text = retPhoto(grid1.TextMatrix(i, 2))
        tCard!isPhoto = True
        tCard!CardNo = nCard
        tCard.Update
        
'        If mRound(grid1.TextMatrix(i, 6)) > 1 Then
'            If validPhoto(retPhoto(grid1.TextMatrix(i, 0)), True) Then
'                '«·’Ê—… «·’€Ì—…
'                tCard.AddNew
'                tCard!Right = MyMeasure(5.1) + MyMeasure(0.25)
'                tCard!Top = MyMeasure(2) + MyMeasure(3) - MyMeasure(1.2) - MyMeasure(0.6) - MyMeasure(0.4) - MyMeasure(0.15) + MyMeasure(0.17)
'                tCard!Width = MyMeasure(1.05) * 0.95
'                tCard!Height = MyMeasure(1.2) * 0.95
'                tCard!text = TurnValue(retPhoto(grid1.TextMatrix(i, 0)), "", Null)
'                tCard!isPhoto = True
'                tCard!CardNo = nCard
'                tCard.Update
'            End If
'        End If
    End If
Next
prog1.Visible = False
tCard.Requery
doprint2 = Not (tCard.EOF And tCard.BOF)
Set CardTable = Nothing
End With
End Function
Sub myProc()
If ActiveControl.Name = cmdYear.Name Then
    cmdYear.Tag = oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 0)
    cmdYear.Caption = ArbString(oSearchYear.grid1.TextMatrix(oSearchYear.grid1.Row, 1))
    Unload oSearchYear
ElseIf ActiveControl.Name = cmdYearPaid.Name Then
    cmdYearPaid.Tag = oSearchYearPaid.grid1.TextMatrix(oSearchYearPaid.grid1.Row, 0)
    cmdYearPaid.Caption = ArbString(oSearchYearPaid.grid1.TextMatrix(oSearchYearPaid.grid1.Row, 1))
    Unload oSearchYearPaid
Else
    ActiveControl.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    xcode_desca.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
    Unload oSearch
End If
End Sub
Private Sub checkPhoto()
Dim aPrint As Variant
With grid1
prog1.Value = 0
prog1.Visible = True
For i = 1 To grid1.rows - 1
    prog1.Value = Round(i / (grid1.rows - 1), 2) * 100
    aPrint = Printed(.TextMatrix(i, 0), .TextMatrix(i, 1), cmdYear.Tag, con)
    grid1.TextMatrix(i, 11) = myFormat_p(retFlag(aPrint, "date"))
    If IsDate(grid1.TextMatrix(i, 11)) Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE0E0E0
    Else
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
    End If
    If Not validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then
        grid1.Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
    ElseIf Val(grid1.TextMatrix(i, grid1.Cols - 2)) <> 0 Then
        grid1.Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbGreen
    End If
Next
prog1.Visible = False
End With
End Sub
Private Sub myLoadGrd()
Dim loctable As ADODB.Recordset, cWhere As String, cString As String, sdate_print As String
Dim aPaid As Variant, aPrint As Variant
Dim nRecordcount As Long, i As Long, bAddRow As Boolean
Me.MousePointer = 11

cString = "SELECT  FILE1_10.*,vw_last_card.doc_no,vw_last_card.date as last_date,FILE6_20H.FORM_NO as last_doc,PAID_TYPES.desca AS TYPE_DESCA" & _
          " FROM FILE1_10 INNER JOIN vw_last_card ON FILE1_10.CODE = vw_last_card.CODE" & _
          " INNER JOIN FILE6_20H ON vw_last_card.DOC_NO = FILE6_20H.DOC_NO" & _
          " INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE"

If chkMobil.Value = 1 Then
    cString = cString & " WHERE (FILE1_10.DIED = 0 OR FILE1_10.CODE IN (SELECT MEMBER FROM FILE1_11)) AND (NOT MOBIL IS NULL)"
Else
    cString = cString & " WHERE FILE1_10.DIED = 0"
End If

If cmdYearPaid.Tag <> "" Then
    cWhere = cWhere & Tr(cWhere) & "vw_last_card.year_code >= " & cmdYearPaid.Tag
End If

If ValidInt(xCode1.text) Then
    cWhere = cWhere & Tr(cWhere) & " File1_10.CODE  " & IIf(IsNumeric(xCode2.text), " >= ", " = ") & xCode1.text
End If

If ValidInt(xCode2.text) Then
    cWhere = cWhere & Tr(cWhere) & " File1_10.CODE <= " & xCode2.text
End If

If xCompany.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & "FILE1_10.COMPANY = " & addvalue(xCompany.BoundText)
End If

If IsDate(xDate1.text) Then
     cWhere = cWhere & Tr(cWhere) & "vw_last_card.Date >= " & DateSq(xDate1.text)
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & Tr(cWhere) & "vw_last_card.DATE <= " & DateSq(xDate2.text)
End If

If cWhere <> "" Then
    cString = cString & " AND " & cWhere
End If

cString = cString & " ORDER BY FILE1_10.DESCA"

Set loctable = New ADODB.Recordset
With grid1
Set loctable = myCmd(cString, con)
If Not (loctable.EOF And loctable.BOF) Then
    nRecordcount = loctable.RecordCount
End If


prog1.Visible = True
prog1.Value = 0
Dim nFound As Long
Dim bGo As Boolean
Dim sCaption As String
sCaption = Me.Caption
    
Do Until loctable.EOF
    i = i + 1
    Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & nRecordcount
    prog1.Value = Round(i / (nRecordcount), 2) * 100
    If Check2.Value = 1 Then
        bGo = True
    Else
        bGo = .FindRow(loctable!code, , 0) = -1
    End If
    If xPrinted.Value = 1 And bGo Then
        aPrint = Printed(loctable!code, "", cmdYear.Tag, con)
        bGo = IsEmpty(aPrint)
    End If
    
    If bGo Then
        .AddItem ""
        ' «·„⁄—Ê÷
        .TextMatrix(grid1.rows - 1, 0) = loctable!code
        .TextMatrix(grid1.rows - 1, 1) = ""
        
        .TextMatrix(grid1.rows - 1, 2) = loctable!code
        .TextMatrix(grid1.rows - 1, 3) = loctable!Title & ""
        .TextMatrix(grid1.rows - 1, 4) = loctable!Desca & ""
        .TextMatrix(grid1.rows - 1, 5) = "⁄÷Ê ⁄«„·"
        .TextMatrix(grid1.rows - 1, 6) = loctable!last_doc & ""
        .TextMatrix(grid1.rows - 1, 7) = myFormat_p(loctable!last_date)
        .TextMatrix(grid1.rows - 1, 8) = loctable!TYPE_desca & ""
        .TextMatrix(grid1.rows - 1, 9) = loctable!Mobil & ""
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing

If ckhPrintMain.Value = 0 And chkMobil.Value = 0 Then
    cString = "SELECT file1_11.Code,File1_11.Relation,File1_11.Member,File1_11.Title,FILE1_10.MOBIL,file1_11.descA,FILE1_10.DESCA as Desca_Member,Relation_codes.Desca as relation_Desca,FILE6_20H.FORM_NO as last_doc,FILE6_20H.DATE as last_date,FILE1_11.GENDER,FILE1_11.RELATION,FILE1_11.DATE_BIRTH,PAID_TYPES.DESCA AS TYPE_DESCA " & _
              " From File1_11 Inner Join File1_10 On File1_11.Member = File1_10.Code " & _
              " INNER join relation_codes on file1_11.relation = relation_codes.code" & _
              " inner join vw_last_card on file1_11.member = vw_last_card.code" & _
              " INNER JOIN FILE6_20H ON vw_last_card.DOC_NO = FILE6_20H.DOC_NO" & _
              " INNER JOIN PAID_TYPES ON FILE6_20H.TYPE = PAID_TYPES.CODE"

    cWhere = cWhere & Tr(cWhere) & "dbo.fn_overAge_id(file1_11.id," & dateSql(sDate_Season) & ",25,default,default,default,default) = 0"
    
    
    If ValidNum(xAppend.text) And ValidNum(xCode1.text) Then
        cWhere = cWhere & turn(cWhere, " and ") & "File1_11.MEMBER = " & xCode1.text & " and File1_11.code = " & xAppend.text
    End If
    
    If cWhere <> "" Then cString = cString & " Where " & cWhere
    
    Set loctable = New ADODB.Recordset
    Set loctable = myCmd(cString, con)
    If Not (loctable.EOF And loctable.BOF) Then
        nRecordcount = loctable.RecordCount
    End If
    prog1.Visible = True
    prog1.Value = 0
    Dim bSkip As Boolean
    i = 0
    Do Until loctable.EOF
        i = i + 1
        Me.Caption = sCaption & " - " & "”Ã· " & i & " „‰ " & nRecordcount
        
        prog1.Value = Round(i / (nRecordcount), 2) * 100
        bSkip = False
       
        If Check2.Value = 1 Then
            bGo = True
        Else
            bGo = .FindRow(loctable!member & "-" & loctable!code, , 2) = -1
        End If

        If xPrinted.Value = 1 And bGo Then
            aPrint = Printed(loctable!member, loctable!code, cmdYear.Tag, con)
            bGo = IsEmpty(aPrint)
        End If
        
        If bGo Then
            .AddItem ""
             ' «·„⁄—Ê÷
            .TextMatrix(.rows - 1, 0) = loctable!member
            .TextMatrix(.rows - 1, 1) = loctable!code
            
            .TextMatrix(.rows - 1, 2) = loctable!member & "-" & loctable!code
            .TextMatrix(.rows - 1, 3) = loctable!Title & ""
            .TextMatrix(.rows - 1, 4) = loctable!Desca & ""
            If loctable!RELATION = 1 Then
                .TextMatrix(grid1.rows - 1, 5) = "⁄÷Ê ⁄«„·"
            ElseIf loctable!RELATION = 2 Then
                .TextMatrix(grid1.rows - 1, 5) = "⁄÷Ê  «»⁄"
            Else
                .TextMatrix(grid1.rows - 1, 5) = "⁄÷Ê  «»⁄"
            End If
                                
            .TextMatrix(grid1.rows - 1, 6) = loctable!last_doc & ""
            .TextMatrix(grid1.rows - 1, 7) = myFormat_p(loctable!last_date)
            .TextMatrix(grid1.rows - 1, 8) = loctable!TYPE_desca & ""
            .TextMatrix(grid1.rows - 1, 9) = loctable!Mobil & ""
        End If
        loctable.MoveNext
    Loop
    loctable.Close
    Set loctable = Nothing
    prog1.Visible = False
End If

Me.Caption = sCaption

Me.MousePointer = 0
If grid1.rows > 1 Then
    grid1.Select 1, 0, 1, 1
    grid1.Sort = flexSortGenericAscending
End If
End With
End Sub
Private Sub Fixgrd()
With grid1

    .TextMatrix(0, 0) = "—ﬁ„ «·⁄÷Ê"
    .TextMatrix(0, 1) = "—ﬁ„ «· «»⁄"
    .TextMatrix(0, 2) = "—ﬁ„ «·⁄÷ÊÌ…"
    .TextMatrix(0, 3) = "«··ﬁ»"
    .TextMatrix(0, 4) = "«·«”„"
    .TextMatrix(0, 5) = "‰Ê⁄ «·⁄÷ÊÌ…"
    .TextMatrix(0, 6) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 7) = " «—ÌŒ «·„” ‰œ"
    .TextMatrix(0, 8) = "‰Ê⁄ «·„” ‰œ"
    .TextMatrix(0, 9) = "—ﬁ„ «·„Õ„Ê·"
                
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
                
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .ColWidth(4) = 4000
    .ColWidth(5) = 2000
    .ColWidth(6) = 1200
    .ColWidth(7) = 1400
    .ColWidth(8) = 1800
    .ColWidth(9) = 1800
    .ColHidden(0) = True
    .ColHidden(1) = True
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColDataType(0) = flexDTLong
    .ColDataType(1) = flexDTLong
    .ExplorerBar = flexExSortShow
End With
End Sub
Private Sub FixgrdExcel()
With grdExcel
.TextMatrix(0, 0) = "#"
.TextMatrix(0, 1) = "«··ﬁ»"

.TextMatrix(0, 2) = "«·≈”„"
.TextMatrix(0, 3) = "‰Ê⁄ «·⁄÷ÊÌ…"

.TextMatrix(0, 4) = "—ﬁ„ «·⁄÷ÊÌ…"
.TextMatrix(0, 5) = "«·’Ê—…"

    
.ColDataType(0) = flexDTDouble
.ColWidth(0) = 1000
.ColWidth(1) = 2000
.ColWidth(2) = 3000
.ColWidth(3) = 2000
.ColWidth(4) = 1000
.ColWidth(5) = 1200
End With
End Sub
Private Sub CalcTotals()
Dim nAll As Long, nPhoto As Long, nPhoto2 As Long, nPages As Long, nrest As Long
StatusBar1.Panels(3).text = ""
StatusBar1.Panels(2).text = ""
StatusBar1.Panels(1).text = ""
If grid1.rows = 1 Then Exit Sub
For i = 0 To grid1.rows - 1
    nAll = nAll + 1
    If validPhoto(retPhoto(grid1.TextMatrix(i, 2))) Then nPhoto = nPhoto + 1
Next
nPhoto2 = ((Val(xRow.text) - 1) * 2) + (Val(xCol.text) - 1)
nPages = Fix(nPhoto / 10)
If nPhoto > 10 Then nLeft = nPhoto Mod 10
StatusBar1.Panels(3).text = "⁄œœ «·”Ã·«  : " & nAll
StatusBar1.Panels(2).text = "⁄œœ «·”Ã·«  »’Ê— : " & nPhoto
StatusBar1.Panels(1).text = "⁄œœ «·’›Õ«  : " & nPages
If nrest > 0 Then StatusBar1.Panels(3).text = StatusBar1.Panels(3).text & turn(StatusBar1.Panels(3).text, " ") & nrest & " ’Ê—…"
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
End Sub

Private Sub xDate1_DblClick()
Set datefrm.oDate = xDate1
datefrm.Show 1
End Sub

Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
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
If Not ValidInt(xCode1.text) Then Exit Sub
Dim aRet As Variant
aRet = GetFields("select DESCA from file1_10 where code = " & xCode1.text)
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
Private Function myRev(pString As String) As String
Dim aString As Variant
If Trim(pString) = "" Then Exit Function
aString = Split(Trim(pString), "-")
If UBound(aString) = 0 Then
    myRev = Trim(aString(0))
Else
    myRev = Trim(aString(1)) & "-" & Trim(aString(0))
End If
End Function
Private Function DeletePrint() As Long
Dim nRecords As Long, Row As Long, nRecord As Long
On Error GoTo myerror
prog1.Visible = True
prog1.Value = 0
For Row = 1 To grid1.rows - 1
    prog1.Value = mRound(Row / (grid1.rows - 1), 2) * 100
    con.Execute "delete from file4_10 where member = " & addvalue(grid1.TextMatrix(Row, 0)) & " And " & IIf(grid1.TextMatrix(Row, 1) = "", " file4_10.code is null ", "file4_10.code = " & addvalue(grid1.TextMatrix(Row, 1))) & " AND [YEAR] = " & addvalue(cmdYear.Tag), nRecord
    nRecords = nRecords + nRecord
Next
DeletePrint = nRecords
prog1.Visible = False
prog1.Value = 0
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
