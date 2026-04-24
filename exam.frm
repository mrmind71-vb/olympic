VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form examfrm 
   Caption         =   "„ﬁ«” «·‰Ÿ«—« "
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   0
      Width           =   7485
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   6255
         Picture         =   "exam.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   5220
         Picture         =   "exam.frx":242A
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "exam.frx":4BFD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "exam.frx":71A9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "exam.frx":9A43
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "exam.frx":BEAF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave 
         Height          =   510
         Left            =   3150
         MaskColor       =   &H00FFFFFF&
         Picture         =   "exam.frx":E428
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   5580
      Width           =   3210
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   47
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
         Picture         =   "exam.frx":1078B
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "exam.frx":1295B
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   48
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
         Picture         =   "exam.frx":14AA3
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "exam.frx":16C6B
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1605
         TabIndex        =   49
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
         Picture         =   "exam.frx":18DBA
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "exam.frx":1AF9A
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2385
         TabIndex        =   50
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
         Picture         =   "exam.frx":1D0F5
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "exam.frx":1F2B1
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1725
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   1395
      Width           =   2220
      Begin VB.CheckBox xAdd2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·Ê«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   540
         Width           =   1590
      End
      Begin VB.CheckBox xAdd4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "›Ê Ê Ã—«Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1260
         Width           =   1590
      End
      Begin VB.CheckBox xAdd3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "›Ê Ê »—Ê«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   900
         Width           =   1590
      End
      Begin VB.CheckBox xAdd1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«‰ Ï ›·«·ﬂ‘‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   1590
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1005
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   675
      Width           =   5460
      Begin VB.TextBox xDesca 
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
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "DATE"
         Top             =   585
         Width           =   3840
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   630
         Width           =   495
      End
      Begin VB.Label xCode 
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ :"
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
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2040
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1665
      Width           =   5460
      Begin VB.TextBox xTotal 
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
         Left            =   1575
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Tag             =   "DATE"
         Top             =   1620
         Width           =   2445
      End
      Begin VB.TextBox xdate 
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
         Left            =   1575
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "DATE"
         Top             =   900
         Width           =   2445
      End
      Begin VB.CommandButton cmdType1 
         Caption         =   "..."
         Height          =   330
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   330
      End
      Begin VB.CommandButton cmdType2 
         Caption         =   "..."
         Height          =   330
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   330
      End
      Begin VB.TextBox xdateDelivery 
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
         Left            =   1575
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "DATE"
         Top             =   1260
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xType1 
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
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
      Begin MSDataListLib.DataCombo xType2 
         Height          =   315
         Left            =   1575
         TabIndex        =   3
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·ﬂ‘› :"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   945
         Width           =   630
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì⁄«œ «· ”·Ì„ :"
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
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1305
         Width           =   1050
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·⁄œ”«  :"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·‘„»— : "
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   1020
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
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9360
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   5895
      Top             =   6435
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Left            =   5940
      Top             =   5310
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Left            =   4545
      Top             =   6435
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Left            =   -1575
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
   Begin VB.Frame Frame2 
      Caption         =   "Left Eye"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   4050
      TabIndex        =   28
      Top             =   3735
      Width           =   3840
      Begin VB.TextBox xAxis_Left2 
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
         Left            =   2655
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1305
         Width           =   960
      End
      Begin VB.TextBox xCyl_Left2 
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
         Left            =   1440
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1305
         Width           =   960
      End
      Begin VB.TextBox xSph_left2 
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
         Left            =   225
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1305
         Width           =   960
      End
      Begin VB.TextBox xAxis_left1 
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
         Left            =   2655
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox xCyl_left1 
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
         Left            =   1440
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox xSph_left1 
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
         Left            =   225
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2925
         TabIndex        =   31
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cyl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1665
         TabIndex        =   30
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sph."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   405
         TabIndex        =   29
         Top             =   360
         Width           =   435
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   2520
         Y1              =   270
         Y2              =   1755
      End
      Begin VB.Line Line4 
         X1              =   135
         X2              =   3735
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line2 
         X1              =   1305
         X2              =   1305
         Y1              =   270
         Y2              =   1755
      End
      Begin VB.Line Line1 
         X1              =   135
         X2              =   3735
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Shape Shape1 
         Height          =   1500
         Left            =   135
         Top             =   270
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Right Eye"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   180
      TabIndex        =   32
      Top             =   3735
      Width           =   3840
      Begin VB.TextBox xsph_right1 
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
         Left            =   225
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox xCyl_right1 
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
         Left            =   1485
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox xAxis_right1 
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
         Left            =   2700
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox xsph_right2 
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
         Left            =   225
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox xCyl_right2 
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
         Left            =   1485
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox xAxis_Right2 
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
         Left            =   2700
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1350
         Width           =   960
      End
      Begin VB.Shape Shape2 
         Height          =   1500
         Left            =   90
         Top             =   270
         Width           =   3660
      End
      Begin VB.Line Line8 
         X1              =   90
         X2              =   3735
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line7 
         X1              =   1260
         X2              =   1260
         Y1              =   270
         Y2              =   1755
      End
      Begin VB.Line Line6 
         X1              =   90
         X2              =   3735
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line5 
         X1              =   2520
         X2              =   2520
         Y1              =   270
         Y2              =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sph."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   405
         TabIndex        =   35
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cyl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1665
         TabIndex        =   34
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Axis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2925
         TabIndex        =   33
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame7 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   3105
      Width           =   2220
      Begin VB.TextBox xLpd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   855
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.P.D :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   37
         Top             =   225
         Width           =   630
      End
   End
End
Attribute VB_Name = "examfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Public sDoc_No As String, sCode As String, myForm As Form
Public bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As New ADODB.Recordset
Dim oSearchDoc As New Search3
Const LoadMode = 1, DefineMode = 2

Private Sub CmdInform_Click()
CardLookup
End Sub

Private Sub CmdPrint_Click()
doprint
End Sub
Private Sub cmdType1_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sSave = xType1.BoundText
oFlagfrm.sTable = "type_codes1"
oFlagfrm.sCaption = "«‰Ê«⁄ «·‘‰«»—"
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data1.Refresh
xType1.BoundText = sCode
If Not xType1.MatchedWithList Then xType1.BoundText = ""
End Sub
Private Sub cmdType2_Click()
Dim oFlagfrm As New flag_mainfrm, sCode As String
sSave = xType1.BoundText
oFlagfrm.sTable = "type_codes2"
oFlagfrm.sCaption = "«‰Ê«⁄ «·⁄œ”« "
oFlagfrm.nZero = -1
oFlagfrm.bEdit = True
oFlagfrm.Show 1
data2.Refresh
xType2.BoundText = sCode
If Not xType2.MatchedWithList Then xType2.BoundText = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM TYPE_CODES1"
Set xType1.RowSource = data1
xType1.ListField = "Desca"
xType1.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM TYPE_CODES2"
Set xType2.RowSource = data2
xType2.ListField = "Desca"
xType2.BoundColumn = "Code"
openCardTable
If sCode = "" Then mydefine Else myUndo
End Sub
Private Sub CmdAdd_Click()
mydefine
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
    con.BeginTrans
    con.Execute "Delete From Exam  Where code = " & xCode.Caption
    con.Execute "UPDATE FILE6_20H SET FILE6_20H.TOTAL_EXAM = dbo.f_inv_exam(FILE6_20H.DOC_NO) WHERE DOC_NO = " & MyParn(sDoc_No)
    con.CommitTrans
    openCardTable
    myUndo
    Unload Me
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
Inform " „ Õ›Ÿ »Ì«‰«  «·ﬂ‘› »‰Ã«Õ"
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xCode.Tag = DefineMode Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & xCode.Caption, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode) And bEdit
CmdDel.Enabled = (nMode = LoadMode) And bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
cmdSave.Enabled = bEdit
xCode.Tag = nMode
End Sub
Sub mydefine()
xCode.Caption = Newflag("EXAM", "CODE")
xDesca.Text = ""
xType1.BoundText = ""
xType2.BoundText = ""
xdateDelivery.Text = ""
xTotal.Text = ""
xdate.Text = ""
xAxis_left1.Text = ""
xCyl_left1.Text = ""
xSph_left1.Text = ""
xAxis_Left2.Text = ""
xCyl_Left2.Text = ""
xSph_left2.Text = ""
xAxis_right1.Text = ""
xCyl_right1.Text = ""
xsph_right1.Text = ""
xAxis_Right2.Text = ""
xCyl_right2.Text = ""
xsph_right2.Text = ""
xLpd.Text = ""
Handlecontrols DefineMode
End Sub
Private Sub myload()
xCode.Caption = CardTable!CODE & ""
xDesca.Text = CardTable!Desca & ""
xType1.BoundText = CardTable!type1 & ""
xType2.BoundText = CardTable!type2 & ""
xdateDelivery.Text = Format(CardTable!dateDelivery, "dd-mm-yyyy")
xTotal.Text = Myvalue(CardTable!TOTAL & "")
xdate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xAxis_left1.Text = CardTable!Axis_left1 & ""
xCyl_left1.Text = CardTable!Cyl_left1 & ""
xSph_left1.Text = CardTable!Sph_left1 & ""
xAxis_Left2.Text = CardTable!AxIS_LEFT2 & ""
xCyl_Left2.Text = CardTable!Cyl_Left2 & ""
xSph_left2.Text = CardTable!Sph_left2 & ""
xAxis_right1.Text = CardTable!Axis_Right1 & ""
xCyl_right1.Text = CardTable!Cyl_right1 & ""
xsph_right1.Text = CardTable!sph_right1 & ""
xAxis_Right2.Text = CardTable!Axis_Right2 & ""
xCyl_right2.Text = CardTable!Cyl_right2 & ""
xsph_right2.Text = CardTable!sph_right2 & ""
xLpd.Text = CardTable!Lpd & ""
xAdd1.Value = IIf(CardTable!ADD1, 1, 0)
xAdd2.Value = IIf(CardTable!ADD2, 1, 0)
xAdd3.Value = IIf(CardTable!ADD3, 1, 0)
xAdd4.Value = IIf(CardTable!ADD4, 1, 0)
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "DOC_NO", addstring(sDoc_No))
aInsert = AddFlag(aInsert, "DESCA", addstring(xDesca.Text))
aInsert = AddFlag(aInsert, "TYPE1", addvalue(xType1.BoundText))
aInsert = AddFlag(aInsert, "TYPE2", addvalue(xType2.BoundText))
aInsert = AddFlag(aInsert, "[Date]", addDate(xdate.Text))
aInsert = AddFlag(aInsert, "DateDelivery", addDate(xdateDelivery.Text))
aInsert = AddFlag(aInsert, "TOTAL", Val(xTotal.Text))
aInsert = AddFlag(aInsert, "[Axis_left1]", addstring(xAxis_left1.Text))
aInsert = AddFlag(aInsert, "[Axis_left2]", addstring(xAxis_Left2.Text))
aInsert = AddFlag(aInsert, "[Cyl_left1]", addstring(xCyl_left1.Text))
aInsert = AddFlag(aInsert, "[Cyl_left2]", addstring(xCyl_Left2.Text))
aInsert = AddFlag(aInsert, "[sph_left1]", addstring(xSph_left1.Text))
aInsert = AddFlag(aInsert, "[sph_left2]", addstring(xSph_left2.Text))
aInsert = AddFlag(aInsert, "[Axis_right1]", addstring(xAxis_right1.Text))
aInsert = AddFlag(aInsert, "[Axis_right2]", addstring(xAxis_Right2.Text))
aInsert = AddFlag(aInsert, "[Cyl_right1]", addstring(xCyl_right1.Text))
aInsert = AddFlag(aInsert, "[Cyl_right2]", addstring(xCyl_right2.Text))
aInsert = AddFlag(aInsert, "[sph_right1]", addstring(xsph_right1.Text))
aInsert = AddFlag(aInsert, "[sph_right2]", addstring(xsph_right2.Text))
aInsert = AddFlag(aInsert, "[Lpd]", addstring(xLpd.Text))
aInsert = AddFlag(aInsert, "[ADD1]", xAdd1.Value)
aInsert = AddFlag(aInsert, "[ADD2]", xAdd2.Value)
aInsert = AddFlag(aInsert, "[ADD3]", xAdd3.Value)
aInsert = AddFlag(aInsert, "[ADD4]", xAdd4.Value)
On Error GoTo myerror
con.BeginTrans
If xCode.Tag = DefineMode Then
    xCode.Caption = Newflag("EXAM", "CODE")
    aInsert = AddFlag(aInsert, "CODE", xCode.Caption)
    con.Execute addInsert(aInsert, "EXAM")
Else
    con.Execute addUpdate(aInsert, "EXAM", "EXAM.code = " & xCode.Caption)
End If
con.Execute "UPDATE FILE6_20H SET FILE6_20H.TOTAL_EXAM = dbo.f_inv_exam(FILE6_20H.DOC_NO) WHERE DOC_NO = " & MyParn(sDoc_No)
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myproc()
   CardTable.Find "code = " & oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0), , adSearchForward, adBookmarkFirst
   myload
   oSearchDoc.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetKbLayout Lang_AR
On Error Resume Next
rdcode.Requery
CardTable.Close
Set CardTable = Nothing
Set examfrm = Nothing
Err.Clear
End Sub
Private Sub Option1_Click(Index As Integer)
    'Cmd_Tree_Click
End Sub
Private Sub xCode_LostFocus()
'xCode.BackColor = &H80000005
'If XCODE.CAPTION = "" Then Exit Sub
'CardTable.Find "code = " & XCODE.CAPTION, , adSearchForward, adBookmarkFirst
'If Not CardTable.EOF Then MyLoad
'SetKbLayout Lang_AR
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then Cmd_Tree_Click
End Sub
Private Sub xSection_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If xSection.BoundText = "" Then Exit Sub
    If Not xSection.MatchedWithList Then
        If MsgBox("≈÷«›… „Ã„Ê⁄… ÃœÌœ…", vbYesNo + vbDefaultButton1, "«÷«›… ﬁ”„") = vbYes Then
            On Error Resume Next
            nCode = Newflag("EXAMSC", "code")
            For i = 1 To 10
                con.BeginTrans
                    con.Execute "insert into EXAMSC(code,[desca]) " & _
                    "values(" & _
                    addvalue(nCode) & "," & _
                    addstring(Trim(xSection.Text)) & _
                    ")"
                    If Err.Number = 0 Then Exit For
                    If Err.Number = -2147467259 Then
                        Err.Clear
                        con.RollbackTrans
                        nCode = nCode + 1
                    End If
                    If Err.Number <> 0 Then GoTo myerror
            Next
            con.CommitTrans
            data2.Refresh
            xSection.BoundText = nCode
        End If
    End If
End If
Exit Sub
myerror:
    con.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub


Private Sub xgroup_LostFocus()
If Not xgroup.MatchedWithList Then
    xgroup.BoundText = ""
End If
End Sub
Private Sub xgroup2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And (xGroup2.MatchedWithList Or Trim(xGroup2.BoundText) = "") Then Cmd_Tree_Click
End Sub
Private Sub xgroup2_LostFocus()
If Not xGroup2.MatchedWithList Then xGroup2.BoundText = ""
End Sub
Function MYVALID() As Boolean
If xCode.Caption = "" Then
    MsgBox "ﬂÊœ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDesca.Text = "" Then
    MsgBox "≈”„ «·’‰› ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If
MYVALID = True
End Function
Private Sub xLpd_GotFocus()
myGotFocus xLpd
End Sub
Private Sub xsph_right1_GotFocus()
myGotFocus xsph_right1
End Sub
Private Sub xCyl_right1_GotFocus()
myGotFocus xCyl_right1
End Sub
Private Sub xAxis_Right1_GotFocus()
myGotFocus xAxis_right1
End Sub
Private Sub xsph_right2_GotFocus()
myGotFocus xsph_right2
End Sub
Private Sub xCyl_right2_GotFocus()
myGotFocus xCyl_right2
End Sub
Private Sub xAxis_Right2_GotFocus()
myGotFocus xAxis_Right2
End Sub
Private Sub xAxIS_LEFT2_GotFocus()
myGotFocus xAxis_Left2
End Sub
Private Sub xCyl_Left2_GotFocus()
myGotFocus xCyl_Left2
End Sub
Private Sub xSph_left2_GotFocus()
myGotFocus xSph_left2
End Sub
Private Sub xAxis_left1_GotFocus()
myGotFocus xAxis_left1
End Sub
Private Sub xCyl_left1_GotFocus()
myGotFocus xCyl_left1
End Sub
Private Sub xSph_left1_GotFocus()
myGotFocus xSph_left1
End Sub
Private Sub xLpd_LostFocus()
myLostFocus xLpd
End Sub
Private Sub xsph_right1_LostFocus()
myLostFocus xsph_right1
End Sub
Private Sub xCyl_right1_LostFocus()
myLostFocus xCyl_right1
End Sub
Private Sub xAxis_Right1_LostFocus()
myLostFocus xAxis_right1
End Sub
Private Sub xsph_right2_LostFocus()
myLostFocus xsph_right2
End Sub
Private Sub xCyl_right2_LostFocus()
myLostFocus xCyl_right2
End Sub
Private Sub xAxis_Right2_LostFocus()
myLostFocus xAxis_Right2
End Sub
Private Sub xAxIS_LEFT2_LostFocus()
myLostFocus xAxis_Left2
End Sub
Private Sub xCyl_Left2_LostFocus()
myLostFocus xCyl_Left2
End Sub
Private Sub xSph_left2_LostFocus()
myLostFocus xSph_left2
End Sub
Private Sub xAxis_left1_LostFocus()
myLostFocus xAxis_left1
End Sub
Private Sub xCyl_left1_LostFocus()
myLostFocus xCyl_left1
End Sub
Private Sub xSph_left1_LostFocus()
myLostFocus xSph_left1
End Sub
Private Sub xtotal_GotFocus()
myGotFocus xTotal
End Sub
Private Sub xdateDelivery_GotFocus()
myGotFocus xdateDelivery
End Sub
Private Sub xType1_Validate(Cancel As Boolean)
If Not xType1.MatchedWithList Then xType1.BoundText = ""
End Sub
Private Sub xType2_Validate(Cancel As Boolean)
If Not xType2.MatchedWithList Then xType2.BoundText = ""
End Sub
Private Sub xtype2_GotFocus()
myGotFocus xType2
End Sub
Private Sub xtype1_GotFocus()
myGotFocus xType1
End Sub
Private Sub xCode_GotFocus()
'XCODE.SelStart = 0
'XCODE.SelLength = Len(XCODE.Caption)
'XCODE.BackColor = &HE7E7E7
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xdate_GotFocus()
myGotFocus xdate
End Sub
Private Sub xtotal_LostFocus()
myLostFocus xTotal
End Sub
Private Sub xdateDelivery_LostFocus()
myLostFocus xdateDelivery
End Sub
Private Sub xtype2_LostFocus()
myLostFocus xType2
End Sub
Private Sub xtype1_LostFocus()
myLostFocus xType1
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xDate_LostFocus()
myLostFocus xdate
End Sub
Private Sub xdateDelivery_Validate(Cancel As Boolean)
myValidDate xdateDelivery
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
myValidDate xdate
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code,DescA,CONVERT(VARCHAR(10),[DATE],111),CONVERT(VARCHAR(10),[DATEDELIVERY],111) From EXAM"
Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & "doc_no = " & MyParn(sDoc_No)
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·«”„"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = " «—ÌŒ «· ”·Ì„"
GrdArray(3, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "≈” ⁄·«„ "
oSearchDoc.Show 1
End Sub
Private Function doprint()
On Error GoTo myerror
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With temptable
    temptable.AddNew
    temptable!str26 = "»Ì«‰ ﬂ‘› ··⁄„Ì· : " & GetDesca("Select file3_10.desca from file3_10 inner join file6_20h on file3_10.code = file6_20h.code where file6_20h.doc_no = " & MyParn(sDoc_No))
    temptable!str1 = TurnValue(xCode.Caption)
    temptable!str3 = TurnValue(xDesca.Text)
    temptable!str4 = TurnValue(xType1.Text)
    temptable!str5 = TurnValue(xType2.Text)
    temptable!str6 = TurnValue(xdateDelivery.Text)
    temptable!val1 = Val(xTotal.Text)
    temptable!str8 = TurnValue(xdate.Text)
    temptable!str9 = TurnValue(xMan.Text)

    temptable!str10 = TurnValue(xAxis_left1.Text)
    temptable!Str11 = TurnValue(xCyl_left1.Text)
    temptable!str12 = TurnValue(xSph_left1.Text)
    temptable!str13 = TurnValue(xAxis_Left2.Text)
    temptable!str14 = TurnValue(xCyl_Left2.Text)
    temptable!str15 = TurnValue(xSph_left2.Text)
    temptable!str16 = TurnValue(xAxis_right1.Text)
    temptable!str17 = TurnValue(xCyl_right1.Text)
    temptable!str18 = TurnValue(xsph_right1.Text)
    temptable!str19 = TurnValue(xAxis_Right2.Text)
    temptable!STR20 = TurnValue(xCyl_right2.Text)
    temptable!str21 = TurnValue(xsph_right2.Text)
    temptable!str22 = TurnValue(xLpd.Text)
    
    temptable.Update
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If
contemp.BeginTrans
contemp.CommitTrans
temptable.Requery
main.REPORT1.ReportFileName = App.Path & "\Reports\exam.rpt"
main.REPORT1.DataFiles(0) = tempPath
main.REPORT1.Action = 1
doprint = True
closeCon:
temptable.Close
Set temptable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo closeCon
End Function
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM EXAM"
If sDoc_No <> "" Then cString = cString & turn(cString) & " DOC_NO = " & MyParn(sDoc_No)
cString = cString & " Order by EXAM.CODE"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xCode.Caption <> "" Then
        CardTable.Find "code = " & xCode.Caption, , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

