VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form maindoorfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÇÓÊÚáÇã ÇáÇÚÖÇÁ"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main_door.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   19290
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      ScaleHeight     =   660
      ScaleWidth      =   4935
      TabIndex        =   49
      Top             =   9720
      Width           =   4965
      Begin VB.TextBox xDrive 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   135
         Width           =   555
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   510
         Left            =   1395
         TabIndex        =   52
         Top             =   90
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "main_door.frx":12632
         Caption         =   "ÓÍÈ ÇáÈíÇäÇÊ "
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "main_door.frx":15027
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   3195
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   45
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "main_door.frx":178C0
         Caption         =   "ÎÑæÌ"
         ButtonStyle     =   2
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "main_door.frx":19C7E
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drive "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   8460
      RightToLeft     =   -1  'True
      ScaleHeight     =   615
      ScaleWidth      =   4800
      TabIndex        =   47
      Top             =   8505
      Width           =   4830
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   2430
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   45
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "main_door.frx":1C5DA
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÇÓÇÓí"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "main_door.frx":1E9A5
      End
      Begin Threed.SSCommand cmdInformRel 
         Height          =   510
         Left            =   45
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   45
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "main_door.frx":20A4E
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÊÇÈÚ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "main_door.frx":22E19
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8385
      Left            =   8460
      RightToLeft     =   -1  'True
      ScaleHeight     =   8355
      ScaleWidth      =   4800
      TabIndex        =   30
      Top             =   90
      Width           =   4830
      Begin VB.PictureBox fmDirect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   90
         RightToLeft     =   -1  'True
         ScaleHeight     =   915
         ScaleWidth      =   4245
         TabIndex        =   41
         Top             =   7425
         Width           =   4245
         Begin Threed.SSCommand cmdFirst 
            Default         =   -1  'True
            Height          =   420
            Left            =   2835
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   0
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
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
            Picture         =   "main_door.frx":24EC2
            Caption         =   "Ãæá"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   1
            PictureDisabledFrames=   1
            PictureDisabled =   "main_door.frx":27069
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   1935
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   0
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
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
            Picture         =   "main_door.frx":290B0
            Caption         =   "ÓÇÈÞ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   1
            PictureDisabledFrames=   1
            PictureDisabled =   "main_door.frx":2B19B
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   990
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   0
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
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
            Picture         =   "main_door.frx":2D195
            Caption         =   "ÊÇáí"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   1
            PictureDisabledFrames=   1
            PictureDisabled =   "main_door.frx":2F2A6
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   0
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   741
            _Version        =   196610
            BackColor       =   16777215
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
            Picture         =   "main_door.frx":312A0
            Caption         =   "ÇÎíÑ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   1
            PictureDisabledFrames=   1
            PictureDisabled =   "main_door.frx":334C4
         End
         Begin VB.Label xRecord_No 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   450
            Width           =   3750
         End
      End
      Begin ImgXCtrl6.ImgXCtrl Photo_main 
         DragIcon        =   "main_door.frx":35595
         Height          =   5685
         Left            =   90
         TabIndex        =   31
         Tag             =   "-1"
         Top             =   90
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   10028
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÞÑÇÈÉ"
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   6660
         Width           =   1005
      End
      Begin VB.Label xRelation_Desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   6615
         Width           =   3795
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "íäÊåí Ýí"
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   7065
         Width           =   1140
      End
      Begin VB.Label xCard_end 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   7020
         Width           =   3795
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÇÓã"
         Height          =   330
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   6300
         Width           =   1005
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   6210
         Width           =   3795
      End
      Begin VB.Label xType_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   5805
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÑÞã "
         Height          =   330
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   5895
         Width           =   1140
      End
      Begin VB.Label xMember 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   5805
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9600
      Left            =   90
      RightToLeft     =   -1  'True
      ScaleHeight     =   9570
      ScaleWidth      =   8310
      TabIndex        =   5
      Top             =   90
      Width           =   8340
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":359D7
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   1
         Left            =   6165
         TabIndex        =   6
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":35E19
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   2
         Left            =   4185
         TabIndex        =   7
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":3625B
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   3
         Left            =   2205
         TabIndex        =   8
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":3669D
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   4
         Left            =   225
         TabIndex        =   9
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":36ADF
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   5
         Left            =   6165
         TabIndex        =   10
         Tag             =   "-1"
         Top             =   3375
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":36F21
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   6
         Left            =   4185
         TabIndex        =   11
         Tag             =   "-1"
         Top             =   3375
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":37363
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   7
         Left            =   2205
         TabIndex        =   12
         Tag             =   "-1"
         Top             =   3375
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":377A5
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   8
         Left            =   225
         TabIndex        =   13
         Tag             =   "-1"
         Top             =   3375
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":37BE7
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   9
         Left            =   6165
         TabIndex        =   14
         Tag             =   "-1"
         Top             =   6660
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":38029
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   10
         Left            =   4185
         TabIndex        =   15
         Tag             =   "-1"
         Top             =   6660
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":3846B
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   11
         Left            =   2205
         TabIndex        =   16
         Tag             =   "-1"
         Top             =   6660
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "main_door.frx":388AD
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   12
         Left            =   225
         TabIndex        =   17
         Tag             =   "-1"
         Top             =   6660
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   2
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   3
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   4
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   5
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   5715
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   6
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   5715
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   7
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   5715
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   8
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   5730
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   9
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   9000
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   10
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   9000
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   11
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   9000
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   12
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   9000
         Width           =   1950
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1005
      Left            =   14985
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5130
      Visible         =   0   'False
      Width           =   4065
      Begin VB.CommandButton cmdGetPhoto 
         Caption         =   "ÓÍÈ ßá ÇáÕæÑ"
         Height          =   450
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetPhotoNew 
         Caption         =   "ÓÍÈ ÇáÕæÑÉ ÇáÍÏíËÉ"
         Height          =   450
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdGetData 
         Caption         =   "ÓÍÈ ÇáÈíÇäÇÊ"
         Height          =   450
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2220
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   10545
      Visible         =   0   'False
      Width           =   19290
      _ExtentX        =   34025
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdType 
      Height          =   510
      Left            =   8460
      TabIndex        =   54
      Top             =   9180
      Visible         =   0   'False
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   900
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
      Caption         =   "äæÚ ÇáÚÖæíÉ"
      ButtonStyle     =   4
   End
   Begin VB.Label XCODE_REL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15345
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   855
      Width           =   2040
   End
   Begin VB.Label xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15345
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   450
      Width           =   2040
   End
End
Attribute VB_Name = "maindoorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conMdb As New ADODB.Connection, oSearchMember As New Search_mdb
Dim cardTable As ADODB.Recordset

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetData_Click()
Dim fs As New FileSystemObject
If Trim(xDrive.Text) = "" Then
    MsgBox "ÇáÞÑÕ ÛíÑ ãÓÌá"
    Exit Sub
End If

Dim sSource As String, sTarget As String
sSource = xDrive.Text & ":\etahad_door_sql\data_trans.mdb"
On Error GoTo myerror
If fs.FileExists(sSource) Then
    CloseData
    fs.CopyFile sSource, App.Path & "\mdb_door\data.mdb"
    OpenData
    GetPhotos
    MsgBox "Êã ÓÍÈ ÇáÈíÇäÇÊ ÈäÌÇÍ"
Else
    MsgBox "ãáÝ ÇáÈíÇäÇÊ ÛíÑ ãæÌæÏ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdGo_Click()
myDefine
myload
'xBarCode.Text = ""
End Sub

Private Sub Command1_Click()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select code,[desca] from file1_10"
Generalarray(2) = "Order by FILE1_10.CODE"
Generalarray(3) = 7000
Generalarray(5) = True

listarray(0, 0) = "ÇáÇÓã-ÑÞã ÇáÚÖæíÉ"
listarray(0, 1) = "(VAL('cFilter') = CODE OR %%DESCA%%)"


GrdArray(0, 0) = "ÑÞã ÇáÚÖæíÉ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÇáÇÓã"
GrdArray(1, 1) = 9000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMember.Caption = "ÇÓÊÚáÇã ÇáÇÚÖÇÁ"
oSearchMember.Show 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdInform_Click()
MemberLookup_I
End Sub

Private Sub Form_Load()
SetKbLayout Lang_AR
'xDrive.Text = RetSetting(xDrive.Name, TempSave(Me))
'OpenData
openConMdb conMdb, App.Path & "\MDB\DATA_TRANS.MDB"
myDefine
End Sub
Private Function openCardTable()
Dim cString As String, cWhere As String
Set cardTable = New ADODB.Recordset
cString = "SELECT MEMBERS_INV.* " & _
           " FROM MEMBERS_INV"

cFilter = ""
cFilter = "MEMBERS_INV.MEMBER = " & xCode.Caption
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
cString = cString & " order by MEMBERS_INV.CODE"

Set cardTable = New ADODB.Recordset
cardTable.Open cString, conMdb, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
If (cardTable.BOF And cardTable.EOF) Then
    myDefine
Else
    If XCODE_REL.Caption <> "" Then
        cardTable.Find "CODE = " & addvalue(XCODE_REL.Caption), , adSearchForward, adBookmarkFirst
        If cardTable.EOF Then cardTable.MoveFirst
    Else
        cardTable.MoveFirst
    End If
    myload
End If
End Sub
Private Sub myload()
Dim acode As Variant, loctable As New ADODB.Recordset, nIndex As Long, sPhoto As String, sPhotoRecord

If Not MYVALID(acode) Then
    On Error Resume Next
    Photo1(nIndex).Import.FromFile MainPath & "\error.jpg"
    Err.Clear
    Exit Sub
End If
'MyLoadMember acode

'If retFlag(acode, "TYPE") = "1" Then
'    MyLoadMember acode
'End If
'    If retFlag(aCode, "CODE") = "" And IsNull(LOCTABLE!code) Then
'        If validPhoto(RetPhoto(LOCTABLE!Member)) Then
'            Photo1(0).Picture = LoadPicture(RetPhoto(LOCTABLE!Member))
'        End If
'    ElseIf validInt(retFlag(aCode, "CODE") & "") And Val(retFlag(aCode, "CODE") & "") = Val(LOCTABLE!code & "") Then
'        If validPhoto(RetPhoto(LOCTABLE!Member & "-" & LOCTABLE!code)) Then
'            Photo1(0).Picture = LoadPicture(RetPhoto(LOCTABLE!Member & "-" & LOCTABLE!code))
'        End If
'    Else
'        nIndex = nIndex + 1
'        sPhoto = RetPhoto(LOCTABLE!Member & turn(LOCTABLE!code & "", "-" & LOCTABLE!code))
'        photo1(nIndex) =
'    End If

End Sub
Function aUnMyCodeBar(sCode) As Variant
Dim nVal1 As Integer, nVal2 As Integer, nValSub1 As Integer, nValSub2 As Integer
Dim nNumber1 As String, nNumber2 As String, nNumber3 As String


If Trim(sCode) = "" Then Exit Function
Dim aret As Variant, aSplit As Variant
aSplit = Split(sCode, "-")
If IsEmpty(aSplit) Then Exit Function
If UBound(aSplit) > 1 Then Exit Function
If Not ValidInt(Val(aSplit(0))) Then Exit Function
If UBound(aSplit) = 1 Then
    If Not ValidInt(Val(aSplit(1))) Then Exit Function
End If

'If Not ValidInt(Val(aret(1))) Then Exit Function
'If Not ValidInt(Val(aret(2))) Then Exit Function

'nVal1 = 74: nVal2 = 11: nValSub1 = 71: nValSub2 = 4

'nNumber1 = aret(0)
'nNumber2 = aret(1)
'nNumber3 = aret(2)
'
'nNumber1 = StrReverse(nNumber1)
'nNumber1 = Val(nNumber1) + Val(nVal2)
'nNumber1 = nNumber1 * 2
'nNumber1 = Val(Left(nNumber1, Len(nNumber1) - 1))
'nNumber1 = nNumber1 - nVal1
'
'nNumber2 = Val(nNumber2) + Val(nValSub2)
'nNumber2 = nNumber2 * 2
'nNumber2 = Val(Left(nNumber2, Len(nNumber2) - 1))
'nNumber2 = nNumber2 - nValSub1
'nNumber2 = nNumber2 - Right(nNumber1, 1)
aret = AddFlag(Empty, "MEMBER", aSplit(0))
If UBound(aSplit) = 1 Then aret = AddFlag(aret, "CODE", aSplit(1))
'aret = AddFlag(aret, "TYPE", IIf(Val(nNumber3) = 0, "", Val(nNumber3)))
aUnMyCodeBar = aret
End Function
Private Function MYVALID(acode) As Boolean
If IsEmpty(acode) Then
    n = Beep(1000, 1000)
   MsgBox "ÇáßæÏ ÛíÑ ãæÌæÏ Çæ ÎØÃ Ýì ÇáÈÇÑßæÏ", vbCritical
    Exit Function
End If

If Not ValidInt(retFlag(acode, "MEMBER")) Then
    n = Beep(1000, 1000)
    MsgBox "ÇáßæÏ ÛíÑ ãæÌæÏ Çæ ÎØÃ Ýì ÇáÈÇÑßæÏ", vbCritical
    Exit Function
End If

If (Not ValidInt(retFlag(acode, "CODE"))) And Trim(retFlag(acode, "CODE")) <> "" Then
    n = Beep(1000, 1000)
    MsgBox "ÇáßæÏ ÛíÑ ãæÌæÏ Çæ ÎØÃ Ýì ÇáÈÇÑßæÏ", vbCritical
    Exit Function
End If

If Val(retFlag(acode, "CODE")) > 20 Then
    n = Beep(1000, 1000)
    MsgBox "ÎØÃ Ýì ÇáÈÇÑßæÏ", vbCritical
    Exit Function
End If

'If Not ValidInt(retFlag(acode, "TYPE")) Then
'    n = Beep(1000, 1000)
'    MsgBox "ÎØÃ Ýì äæÚ ÇáÈÇÑßæÏ", vbCritical
'    Exit Function
'End If

If ValidInt(retFlag(acode, "MEMBER")) And Trim(retFlag(acode, "CODE")) = "" Then
    If IsEmpty(GetField("SELECT CODE FROM FILE1_10 WHERE CODE = " & retFlag(acode, "MEMBER"))) Then
        MsgBox "ßæÏ ÇáÚÖæ ÛíÑ ãæÌæÏ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If
End If

If ValidInt(retFlag(acode, "MEMBER")) And ValidInt(retFlag(acode, "CODE")) Then
    If IsEmpty(GetField("SELECT CODE FROM FILE1_10 WHERE CODE = " & retFlag(acode, "MEMBER"))) Then
        MsgBox "ßæÏ ÇáÚÖæ ÇáÇÓÇÓí ÛíÑ ãæÌæÏ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If

    If IsEmpty(GetField("SELECT MEMBER FROM FILE1_11 WHERE MEMBER = " & retFlag(acode, "MEMBER") & " AND CODE = " & retFlag(acode, "CODE"))) Then
        MsgBox "ßæÏ ÇáÚÖæ ÇáÊÇÈÚ ÛíÑ ãæÌæÏ æÇáÇÓÇÓí ãæÌæÏ ", vbCritical
        n = Beep(1000, 1000)
        Exit Function
    End If
End If

MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
closeCon conMdb
addSetting xDrive.Name, xDrive.Text, TempSave(Me)
Set maindoorfrm = Nothing
End Sub

Private Sub Label9_Click()

End Sub

Private Sub Photo1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If Source.Tag <> "" Then
    xBarCode.Text = Source.Tag
    CmdGo_Click
End If
End Sub

Private Sub xBarCode_Change()
cmdGo.Enabled = Trim(xBarCode.Text) <> ""
End Sub
Private Function MyLoadMember(ByVal acode As Variant) As Boolean
Dim aMember As Variant, nCaption As Long
xMember.Caption = retFlag(acode, "MEMBER") & turn(retFlag(acode, "CODE") & "", "-" & retFlag(acode, "CODE"))
xDateLast.Caption = PaidString(retFlag(acode, "MEMBER"))
If retFlag(acode, "CODE") = "" Then
    xdesca.Caption = GetField("select desca from file1_10 where code = " & retFlag(acode, "MEMBER"))
    xType.Caption = "ÇáÚÖæ ÇáÇÓÇÓí"
Else
    Dim aret As Variant
    aret = GetFields("select FILE1_10.DESCA AS MEMBER_DESCA,FILE1_11.desca ,FILE0_00.DESCA AS REL_DESCA from (FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.MEMBER = FILE1_10.CODE) LEFT JOIN FILE0_00 ON (FILE1_11.RELATION = FILE0_00.CODE AND FILE0_00.FLAG = 0) where FILE1_11.member = " & retFlag(acode, "MEMBER") & " and FILE1_11.code = " & retFlag(acode, "code"))
    If Not IsEmpty(aret) Then
        xdesca.Caption = retFlag(aret, "MEMBER_DESCA") & vbCrLf & retFlag(aret, "REL_DESCA") & " " & "(" & retFlag(aret, "DESCA") & ")"
    End If
End If

Dim loctable As New ADODB.Recordset
cString = "SELECT FILE1_10.CODE AS MEMBER,FILE1_10.DESCA,'ÇáÚÖæ äÝÓå' AS REL_DESCA,NULL AS CODE" & _
          " FROM FILE1_10 "
cString = cString & turn(cString) & "FILE1_10.CODE = " & retFlag(acode, "MEMBER")
cString = cString & " UNION ALL "
cString = cString & "SELECT FILE1_10.CODE AS MEMBER,FILE1_11.DESCA ,FILE0_00.DESCA AS REL_DESCA,FILE1_11.CODE" & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.CODE = FILE1_11.MEMBER) LEFT JOIN FILE0_00 ON (FILE1_11.RELATION = FILE0_00.CODE AND FILE0_00.FLAG = 0) "
cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE = " & retFlag(acode, "MEMBER")
cString = cString & turn(cWhere) & cWhere
cString = cString & " ORDER BY CODE"

On Error GoTo myerror
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
sPhoto = retFlag(acode, "MEMBER") & turn(retFlag(acode, "CODE") & "", "-" & retFlag(acode, "CODE"))
Do Until loctable.EOF
    sPhotoRecord = loctable!MEMBER & turn(loctable!code & "", "-" & loctable!code)
    If sPhotoRecord = sPhoto Then
        If validPhoto(RetPhoto(sPhotoRecord)) Then
            Photo1(0).Tag = sPhotoRecord
            Photo1(0).Import.FromFile RetPhoto(sPhoto)
         End If
    Else
        nIndex = nIndex + 1
        If validPhoto(RetPhoto(sPhotoRecord)) Then
            Photo1(nIndex).Visible = True
            Photo1(nIndex).Import.FromFile RetPhoto(sPhotoRecord)
            Photo1(nIndex).Tag = sPhotoRecord
        End If
        xdesca1(nIndex).Caption = loctable!desca & ""
    End If
    nCaption = nCaption + 1
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub myDefine()
Dim i As Long
For i = 1 To Photo1.UBound
    Photo1(i).Images.Clear
    Photo1(i).Tag = ""
    xdesca1(i).Caption = ""
Next
xdesca.Caption = ""
xMember.Caption = ""
xCode.Caption = ""
xType_desca.Caption = ""
xCard_end.Caption = ""
End Sub
Private Sub xBarCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If cmdGo.Enabled Then CmdGo_Click
End If
End Sub
Sub myProc()
xCode.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
oSearchMember.Hide
openCardTable
myUndo
End Sub
Private Sub CloseData()
closeCon con
End Sub
Private Sub GetPhotos()
Dim fs As New FileSystemObject, sSource As String, nRecordCount As Double, i As Long
Dim conMdb As New ADODB.Connection
openConMdb pCon
Dim loctable As ADODB.Recordset
Set loctable = New ADODB.Recordset

loctable.Open "select code as Member,NULL as Serial from file1_10  union all select member,code as Serial from file1_11 ", con, adOpenStatic, adLockReadOnly
If Not loctable.EOF Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
    Prog1.Visible = True
    Prog1.Value = 0
End If
On Error GoTo myerror

Do Until loctable.EOF
    i = i + 1
    bCopy = True
    sCode = loctable!MEMBER & turn(loctable!Serial & "", "-" & loctable!Serial)
    sSource = RetPhotoNew(sCode, , , xDrive & ":\etahad_door")
    sTarget = RetPhoto(sCode)
    
    If fs.FileExists(sSource) Then
        If fs.FileExists(sTarget) Then
           If myFormat(fs.GetFile(sTarget).DateLastModified) >= myFormat(fs.GetFile(sSource).DateLastModified) Then
               bCopy = False
           End If
        End If
        If bCopy Then fs.CopyFile sSource, sTarget
    End If
    Me.Caption = i
    If Prog1.Value <> Int(i / nRecordCount * 100) Then Prog1.Value = Int(i / nRecordCount * 100)
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Prog1.Visible = False
Prog1.Value = 0

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Prog1.Visible = False
Prog1.Value = 0
End Sub

Private Sub xDrive_Change()
xDrive.Text = UCase(xDrive.Text)
End Sub
Private Sub MemberLookup_I(Optional cFilter As String = "", Optional bFilter As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE1_50.CODE,FILE1_50.DESCA,FORMAT(CARD_END,'DD/MM/YYYY')" & _
                  " From FILE1_50 "

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If
Generalarray(2) = "Order by FILE1_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "ÇáÇÓã-ÑÞã ÇáÚÖæ"
listarray(0, 1) = "(%%FILE1_50.DESCA%% OR **FILE1_50.CODE**)"

GrdArray(0, 0) = "ßæÏ ÇáÚÖæ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÇÓã ÇáÚÖæ"
GrdArray(1, 1) = 5500

GrdArray(2, 0) = "ÊÇÑíÎ ÇáÇäÊåÇÁ"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMember.pMdbPath = App.Path & "\MDB\DATA_TRANS.MDB"
oSearchMember.Caption = "ÅÓÊÚáÇã ÇÚÖÇÁ ÇáÏÚæÉ"
oSearchMember.Show 1
End Sub


