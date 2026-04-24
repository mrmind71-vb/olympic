VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form report_insfrm5 
   Caption         =   "»Ì«‰«  «ﬁ”«ÿ „ «Œ—… ⁄‰  «—ÌŒ"
   ClientHeight    =   7230
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11400
   Begin VB.CommandButton cmdDrop 
      Caption         =   "ÿ»«⁄… Œÿ«» «”ﬁ«ÿ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   6075
      Width           =   4515
   End
   Begin VB.CommandButton cmdLetter3 
      Caption         =   "ÿ»«⁄… Œÿ«»«  «‰–«— »œÊ‰ „»«·€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   5445
      Width           =   4515
   End
   Begin VB.Frame Frame3 
      Height          =   3030
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3015
      Width           =   6045
      Begin VB.TextBox xdesca_period 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   225
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Tag             =   "D"
         Top             =   990
         Width           =   4335
      End
      Begin VB.TextBox xMeeting_Date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Tag             =   "D"
         Top             =   630
         Width           =   2985
      End
      Begin VB.TextBox xMeeting_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   2985
      End
      Begin VB.CommandButton cmdLetter2 
         Caption         =   "ÿ»«⁄… Œÿ«»«  «‰ Â«¡ „œ… «·⁄÷ÊÌ… «·„ﬁ”ÿ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2295
         Width           =   5685
      End
      Begin VB.Label Label2 
         Caption         =   "»Ì«‰ «·„Â·…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ «·Ã·”…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·Ã·”…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdLetter 
      Caption         =   "ÿ»«⁄… Œÿ«»«  «‰–«—"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4815
      Width           =   4515
   End
   Begin VB.Frame Frame2 
      Height          =   1680
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   90
      Width           =   4470
      Begin VB.TextBox xValue_install 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1410
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "D"
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox xInstall_Count 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox xValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "≈Ã„«·Ì „‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "⁄‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "⁄œœ «ﬁ”«ÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   585
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "ÿ»«⁄…  ﬁ—Ì— »»Ì«‰«  «·« ’«·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4185
      Width           =   4470
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3600
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1665
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   3600
      Width           =   1500
   End
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3150
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1755
      Width           =   4515
      Begin VB.TextBox xDate_end1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xDate_End2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Tag             =   "D"
         Top             =   585
         Width           =   1410
      End
      Begin VB.TextBox xDate_Begin2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xdate_begin1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   225
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo xStatus 
         Height          =   330
         Left            =   180
         TabIndex        =   8
         Top             =   945
         Width           =   2850
         _ExtentX        =   5027
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
      Begin MSDataListLib.DataCombo xInstall_type 
         Height          =   330
         Left            =   180
         TabIndex        =   9
         Top             =   1305
         Width           =   2850
         _ExtentX        =   5027
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
      Begin VB.Label Label_install 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·”œ«œ"
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
         Height          =   285
         Index           =   6
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Õ«·… «·⁄÷Ê"
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
         Index           =   1
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ ‰Â«Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   " «—ÌŒ »œ«Ì…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   1320
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1485
      Top             =   -450
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Threed.SSCommand cmdStatus 
      Height          =   2265
      Left            =   4680
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   720
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   3995
      _Version        =   196610
      CaptionStyle    =   1
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
      Caption         =   "ﬂ· Õ«·«  «·⁄÷Ê"
      TagVariant      =   "ﬂ· Õ«·«  «·⁄÷Ê"
      ButtonStyle     =   4
   End
End
Attribute VB_Name = "report_insfrm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchYear As New Search_empty, oSearchRegion As New Search
Dim oSearchStatus As New Search_empty
Private Sub cmdApply_Click()
doprint
End Sub
Private Sub cmdDetails_Click()
doprint "REPORT_INS5_2.rpt"
End Sub

Private Sub cmdDrop_Click()
doprint "drop_Ins.rpt", True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Function doprint(Optional pReport As String = "REPORT_INS5.rpt", Optional bAddress As Boolean = False)
Dim sourcetable As New ADODB.Recordset, cOr As String

Dim aHeader(11)
cString = "SELECT file2_10.CODE,FILE2_10.DESCA,FILE2_10.PHONE,FILE2_10.MOBIL,FILE2_10.ADDRESS,dbo.f_last_year_date_install(FILE2_10.CODE) AS LAST_DATE_PAID" & _
          ",INSTALL_CODES.DESCA AS INSTALL_DESCA,FILE2_10.DATE_BEGIN,FILE2_10.DATE_END,FILE2_10.INSTALL_COUNT,FILE2_10.VALUE,FILE2_10.INSTALL_VALUE" & _
          ",SUM(INSTALL_BALANCE.TOTAL_NET - INSTALL_BALANCE.VALUE_PAID) AS REST,Sum(INSTALL_BALANCE.INS_COUNT) AS Ins_Count  " & _
          " FROM File2_10 INNER JOIN INSTALL_BALANCE ON FILE2_10.CODE = INSTALL_BALANCE.CODE " & _
          " LEFT JOIN INSTALL_CODES ON FILE2_10.INSTALL_TYPE = INSTALL_CODES.CODE"
cWhere = "INSTALL_BALANCE.VALUE - INSTALL_BALANCE.VALUE_PAID  > 0"
If bAddress Then cWhere = cWhere & " AND " & "(NOT FILE2_10.ADDRESS IS NULL)"

If IsDate(xDate.text) Then
    aHeader(2) = "«ﬁ”«ÿ „” Õﬁ… „‰ " & myFormat_p(xDate.text)
    cWhere = cWhere & turn(cWhere, " and ") & "INSTALL_BALANCE.DATE_DUE <= " & DateSq(xDate.text)
End If

If xStatus.MatchedWithList Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS = " & addvalue(xStatus.BoundText)
End If

If mRound(xValue_install.text) <> 0 Then
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.INSTALL_VALUE  = " & mRound(xValue_install.text)
End If
If IsDate(xdate_begin1.text) Then
    aHeader(2) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.text), myFormat_p(xDate_Begin2.text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN >= " & DateSq(xdate_begin1.text)
End If

If IsDate(xDate_Begin2.text) Then
    aHeader(2) = " «—ÌŒ »œ«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(myFormat_p(xdate_begin1.text), myFormat_p(xDate_Begin2.text))
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.DATE_BEGIN <= " & DateSq(xDate_Begin2.text)
End If

If IsDate(xDate_end1.text) Then
    aHeader(3) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.text, xDate_End2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end >= " & DateSq(xDate_end1.text)
End If

If IsDate(xDate_End2.text) Then
    aHeader(3) = " «—ÌŒ ‰Â«Ì… «·⁄÷ÊÌ… „‰ " & BetweenString(xDate_end1.text, xDate_End2.text)
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.date_end <= " & DateSq(xDate_End2.text)
End If

If xInstall_type.MatchedWithList Then
    aHeader(4) = "‰Ê⁄ «·«ﬁ”«ÿ : " & xInstall_type.text
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.install_type = " & addvalue(xInstall_type.BoundText)
End If

If cmdStatus.Tag <> "" Then
    aHeader(4) = Replace(cmdStatus.Caption, vbCrLf, "-")
    cWhere = cWhere & turn(cWhere, " and ") & "FILE2_10.STATUS IN(" & cmdStatus.Tag & ")"
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " group by file2_10.CODE,FILE2_10.DESCA,FILE2_10.DATE_BEGIN,FILE2_10.DATE_END,FILE2_10.INSTALL_COUNT,FILE2_10.VALUE,FILE2_10.INSTALL_VALUE,INSTALL_CODES.DESCA,FILE2_10.PHONE,FILE2_10.ADDRESS,FILE2_10.MOBIL"
If mRound(xInstall_Count.text) > 0 Then
    cHaving = "Sum(INSTALL_BALANCE.INS_COUNT) >= " & mRound(xInstall_Count.text)
End If

If mRound(xValue.text) > 0 Then
    cHaving = cHaving & turn(cHaving, " OR ") & " SUM(INSTALL_BALANCE.VALUE - INSTALL_BALANCE.VALUE_PAID) >= " & Val(xValue.text)
End If

If cHaving <> "" Then cString = cString & " HAVING " & cHaving

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext

Dim temptable As New ADODB.Recordset
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With sourcetable
Do Until sourcetable.EOF

    temptable.AddNew
    temptable!val1 = sourcetable!CODE
    temptable!str1 = ArbString(sourcetable!CODE)
    temptable!str2 = sourcetable!Desca
    temptable!Str3 = TurnValue(ArbString(myFormat_p(sourcetable!DATE_END)))
    temptable!Str4 = sourcetable!install_desca
    temptable!str5 = TurnValue(ArbString(myFormat_p(sourcetable!LAST_DATE_PAID)))
    temptable!str6 = TurnValue(ArbString(sourcetable!Address & ""))
    
    If IsNull(sourcetable!Mobil) Then
        temptable!str7 = TurnValue(sourcetable!Phone)
    Else
        temptable!str7 = TurnValue(sourcetable!Mobil)
    End If
    
    temptable!val4 = mRound(sourcetable!install_value)
    temptable!VAL5 = mRound(sourcetable!ins_count)
    temptable!VAL6 = mRound(sourcetable!Rest)
    temptable!str10 = TurnValue(Me.Caption)
    temptable!str11 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str12 = TurnValue(retHeader(aHeader, 3, 5))
    temptable!memo1 = "≈Ã„‹«·Ì «·„»‹·€ «·„” Õ‹ﬁ " & sourcetable!Rest & " Ã‰ÌÂ‹«° ‘«„· «·ﬁÌ„… «·„÷«›… Ê«·€—«„«  «·Ê«Ã»… ·⁄œ„ «·”œ«œ ›Ï „Ê«⁄Ìœ «·«” Õﬁ«ﬁ ÕÌÀ ‰Â«Ì… «·⁄ﬁœ " & ArbString(myFormat_p(sourcetable!DATE_END))
    temptable.Update
    sourcetable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.ReportFileName = sPath_App & "\REPORTS\" & pReport
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set sourcetable = Nothing
End With
End Function

Private Sub cmdLetter_Click()
doprint "Warrent_Ins.rpt", True
End Sub
Private Sub cmdLetter2_Click()
Dim loctable As Recordset
Dim aPrm As Variant

aPrm = AddFlag(aPrm, "address", 1)

If IsDate(xDate.text) Then
    aPrm = AddFlag(aPrm, "date", myFormat_sp(xDate.text))
End If

If xStatus.MatchedWithList Then
    aPrm = AddFlag(aPrm, "status", xStatus.BoundText)
End If

If mRound(xValue_install.text) <> 0 Then
    aPrm = AddFlag(aPrm, "install_value", mRound(xValue_install.text))
End If

If IsDate(xdate_begin1.text) Then
    aPrm = AddFlag(aPrm, "date_begin1", myFormat_sp(xdate_begin1.text))
End If

If IsDate(xDate_Begin2.text) Then
    aPrm = AddFlag(aPrm, "date_begin2", myFormat_sp(xDate_Begin2.text))
End If

If IsDate(xDate_end1.text) Then
    aPrm = AddFlag(aPrm, "date_end1", myFormat_sp(xDate_end1.text))
End If

If IsDate(xDate_End2.text) Then
    aPrm = AddFlag(aPrm, "date_end2", myFormat_sp(xDate_End2.text))
End If

If xInstall_type.MatchedWithList Then
    aPrm = AddFlag(aPrm, "install_type", xInstall_type.BoundText)
End If

If mRound(xInstall_Count.text) > 0 Then
    aPrm = AddFlag(aPrm, "Install_Count", mRound(xInstall_Count.text))
End If

If mRound(xValue.text) > 0 Then
    aPrm = AddFlag(aPrm, "value", mRound(xValue.text))
End If

aPrm = AddFlag(aPrm, "List", cmdStatus.Tag)

Set loctable = myCmd("dbo.sp_late_install", con, adStoredProc, aPrm)
Dim temptable As New ADODB.Recordset
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With loctable
Do Until loctable.EOF
    temptable.AddNew
    temptable!val1 = loctable!CODE
    
    temptable!str1 = ArbString(loctable!CODE)
    temptable!str2 = loctable!Desca
    temptable!Str3 = TurnValue(ArbString(myFormat_p(loctable!DATE_END) & "."))
    temptable!str6 = TurnValue(ArbString(loctable!Address & ""))
    temptable!Str4 = ArbString(vbTab + "Êﬁœ ﬁ—— „Ã·” «·«œ«—… »Ã·”… —ﬁ„ (" & xMeeting_No.text & ") » «—ÌŒ " & myFormat_p(xMeeting_Date.text))
    temptable!Str4 = ArbString(temptable!Str4 & Space(2) & "„‰Õﬂ„ „Â·… ·”œ«œ «·„»«·€ «·„ √Œ—… Ê«·„” Õﬁ… ⁄·Ìﬂ„ ›Ï „Ê⁄œ «ﬁ’«Â " & xdesca_period.text & ".")
    temptable.Update
    loctable.MoveNext
Loop

temptable.Requery
    
If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Report1.ReportFileName = sPath_App & "\REPORTS\WARRENT_INS2.RPT"
    Report1.DataFiles(0) = tempFile
    Report1.Action = 1
End If

Set temptable = Nothing
Set loctable = Nothing
End With
End Sub

Private Sub cmdLetter3_Click()
doprint "Warrent_Ins3.rpt", True
End Sub
Private Sub Form_Load()
openCon con

Set data1.Recordset = myRecordSet("select * from status_Codes", con)
Set xStatus.RowSource = data1
xStatus.ListField = "Desca"
xStatus.BoundColumn = "Code"

Set data3.Recordset = myRecordSet("select * from install_Codes", con)
Set xInstall_type.RowSource = data3
xInstall_type.ListField = "Desca"
xInstall_type.BoundColumn = "Code"

FixRpImage Me

LoadText Me
xStatus.BoundText = 1
End Sub

Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
If Not xType.MatchedWithList Then xType.BoundText = ""
End Sub
Private Sub xJob_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    'Job_Lookup Me, oSearchJob
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
Set report_insfrm5 = Nothing
End Sub

Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xdate_LostFocus()
myLostFocus xDate
myValidDate xDate
End Sub
Private Sub xDate_end1_GotFocus()
myGotFocus xDate_end1
End Sub
Private Sub xDate_end1_LostFocus()
myLostFocus xDate_end1
myValidDate xDate_end1
End Sub
Private Sub xDate_End2_GotFocus()
myGotFocus xDate_End2
End Sub
Private Sub xDate_End2_LostFocus()
myLostFocus xDate_End2
myValidDate xDate_End2
End Sub
Private Sub xdate_begin2_GotFocus()
myGotFocus xDate_Begin2
End Sub
Private Sub xdate_begin2_LostFocus()
myLostFocus xDate_Begin2
myValidDate xDate_Begin2
End Sub
Private Sub xDate_begin1_GotFocus()
myGotFocus xdate_begin1
End Sub
Private Sub xDate_begin1_LostFocus()
myLostFocus xdate_begin1
myValidDate xdate_begin1
End Sub
Private Sub xStatus_GotFocus()
myGotFocus xStatus
End Sub
Private Sub xStatus_LostFocus()
myLostFocus xStatus
If Not xStatus.MatchedWithList Then xStatus.BoundText = ""
End Sub
Private Sub xDate_Period_GotFocus()
myGotFocus xDate_Period
End Sub
Private Sub xDate_Period_LostFocus()
myLostFocus xDate_Period
myValidDate xDate_Period
End Sub
Private Sub xMeeting_date_GotFocus()
myGotFocus xMeeting_Date
End Sub
Private Sub xMeeting_date_LostFocus()
myLostFocus xMeeting_Date
myValidDate xMeeting_Date
End Sub
Private Sub xMeeting_No_GotFocus()
myGotFocus xMeeting_No
End Sub
Private Sub xMeeting_No_LostFocus()
myLostFocus xMeeting_No
End Sub
Private Sub xDesca_Period_GotFocus()
myGotFocus xdesca_period
End Sub
Private Sub xDesca_Period_LostFocus()
myLostFocus xdesca_period
End Sub

Private Sub xValue_install_GotFocus()
myGotFocus xValue_install
End Sub
Private Sub xValue_install_LostFocus()
myLostFocus xValue_install
End Sub

Private Sub xInstall_count_GotFocus()
myGotFocus xInstall_Count
End Sub
Private Sub xInstall_count_LostFocus()
myLostFocus xInstall_Count
End Sub
Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub
Private Sub xInstall_type_GotFocus()
myGotFocus xInstall_type
End Sub
Private Sub xInstall_type_LostFocus()
myLostFocus xInstall_type
If Not xInstall_type.MatchedWithList Then xInstall_type.BoundText = ""
End Sub
Private Sub cmdStatus_Click()
Install_type_Lookup Me, oSearchStatus, , IIf(cmdStatus.Tag <> "", "(CODE NOT IN (" & cmdStatus.Tag & "))", ""), cmdStatus.Tag <> ""
End Sub
Sub myProc()
If oSearchStatus.grid1.TextMatrix(oSearchStatus.grid1.Row, 0) = "" Then
    cmdStatus.Tag = ""
    cmdStatus.Caption = cmdStatus.TagVariant
Else
    If cmdStatus.Tag = "" Then
        cmdStatus.Caption = oSearchStatus.grid1.TextMatrix(oSearchStatus.grid1.Row, 1)
    Else
        cmdStatus.Caption = cmdStatus.Caption & vbCrLf & oSearchStatus.grid1.TextMatrix(oSearchStatus.grid1.Row, 1)
    End If
    cmdStatus.Tag = cmdStatus.Tag & turn(cmdStatus.Tag, ",") & oSearchStatus.grid1.TextMatrix(oSearchStatus.grid1.Row, 0)
End If
oSearchStatus.Hide
End Sub

