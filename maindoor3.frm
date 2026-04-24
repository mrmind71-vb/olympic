VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form maindoorfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÈíÇäÇÊ ÇáÇÚÖÇÁ"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   4  'Icon
   RightToLeft     =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   15990
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrameEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   1335
      ScaleWidth      =   4800
      TabIndex        =   56
      Top             =   6120
      Width           =   4830
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÓäæÇÊ ÛíÑ ãÓÏÏÉ"
         Height          =   330
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label xUnPaid 
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
         TabIndex        =   61
         Top             =   900
         Width           =   3120
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáÓÏÇÏ"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label xLast_date 
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
         TabIndex        =   59
         Top             =   495
         Width           =   3570
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇÎÑ ãæÓã"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label XLAST_PAID 
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
         TabIndex        =   57
         Top             =   90
         Width           =   3570
      End
   End
   Begin VB.PictureBox FrameInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   1335
      ScaleWidth      =   4800
      TabIndex        =   48
      Top             =   6120
      Visible         =   0   'False
      Width           =   4830
      Begin VB.Label xLast_Date_I 
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   90
         Width           =   1590
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇÎÑ ÓÏÇÏ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label xUnPaid_Install 
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   495
         Width           =   3120
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÚÏÏ ÇÞÓÇØ ãÊÃÎÑÉ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   540
         Width           =   1410
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáÞÓØ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label xFirst_Install 
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   900
         Width           =   3120
      End
      Begin VB.Label xInstall_desca 
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   90
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   615
      ScaleWidth      =   4800
      TabIndex        =   41
      Top             =   7515
      Width           =   4830
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   2430
         TabIndex        =   42
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
         Picture         =   "maindoor3.frx":0000
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÇÓÇÓí"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "maindoor3.frx":23CB
      End
      Begin Threed.SSCommand cmdInformRel 
         Height          =   510
         Left            =   45
         TabIndex        =   43
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
         Picture         =   "maindoor3.frx":4474
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÊÇÈÚ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "maindoor3.frx":683F
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   5925
      ScaleWidth      =   4800
      TabIndex        =   31
      Top             =   135
      Width           =   4830
      Begin ImgXCtrl6.ImgXCtrl Photo_main 
         DragIcon        =   "maindoor3.frx":88E8
         Height          =   4065
         Left            =   225
         TabIndex        =   32
         Tag             =   "-1"
         Top             =   135
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   7170
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÝÆÉ ÇáÚÖæíÉ"
         Height          =   330
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   5040
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÇÓã"
         Height          =   330
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   4635
         Width           =   690
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
         TabIndex        =   38
         Top             =   4635
         Width           =   3435
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
         TabIndex        =   37
         Top             =   5040
         Width           =   3435
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÑÞã "
         Height          =   330
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   4275
         Width           =   735
      End
      Begin VB.Label xDate_Birth 
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
         Top             =   5445
         Width           =   3435
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáãíáÇÏ"
         Height          =   330
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   5490
         Width           =   1005
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   4230
         Width           =   2040
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
      TabIndex        =   6
      Top             =   90
      Width           =   8340
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":8D2A
         Height          =   2310
         Index           =   1
         Left            =   6165
         TabIndex        =   7
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":916C
         Height          =   2310
         Index           =   2
         Left            =   4185
         TabIndex        =   8
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":95AE
         Height          =   2310
         Index           =   3
         Left            =   2205
         TabIndex        =   9
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":99F0
         Height          =   2310
         Index           =   4
         Left            =   225
         TabIndex        =   10
         Tag             =   "-1"
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":9E32
         Height          =   2310
         Index           =   5
         Left            =   6165
         TabIndex        =   11
         Tag             =   "-1"
         Top             =   3240
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":A274
         Height          =   2310
         Index           =   6
         Left            =   4185
         TabIndex        =   12
         Tag             =   "-1"
         Top             =   3240
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":A6B6
         Height          =   2310
         Index           =   7
         Left            =   2205
         TabIndex        =   13
         Tag             =   "-1"
         Top             =   3240
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":AAF8
         Height          =   2310
         Index           =   8
         Left            =   225
         TabIndex        =   14
         Tag             =   "-1"
         Top             =   3240
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":AF3A
         Height          =   2310
         Index           =   9
         Left            =   6165
         TabIndex        =   15
         Tag             =   "-1"
         Top             =   6525
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":B37C
         Height          =   2310
         Index           =   10
         Left            =   4185
         TabIndex        =   16
         Tag             =   "-1"
         Top             =   6525
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":B7BE
         Height          =   2310
         Index           =   11
         Left            =   2205
         TabIndex        =   17
         Tag             =   "-1"
         Top             =   6525
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "maindoor3.frx":BC00
         Height          =   2310
         Index           =   12
         Left            =   225
         TabIndex        =   18
         Tag             =   "-1"
         Top             =   6525
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4075
         BackColor       =   16777215
         BorderStyle     =   3
         MousePointer    =   99
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
         Height          =   600
         Index           =   1
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2475
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   4185
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
         Height          =   600
         Index           =   3
         Left            =   2205
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
         Height          =   600
         Index           =   4
         Left            =   225
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
         Height          =   600
         Index           =   5
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   5580
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   6
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   5580
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   7
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   5580
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   8
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   5595
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   9
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   8865
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   10
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   8865
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   11
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   8865
         Width           =   1950
      End
      Begin VB.Label xdesca1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   12
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   8865
         Width           =   1950
      End
   End
   Begin VB.PictureBox fmDirect 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   8550
      RightToLeft     =   -1  'True
      ScaleHeight     =   915
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   8190
      Width           =   4785
      Begin Threed.SSCommand cmdFirst 
         Default         =   -1  'True
         Height          =   420
         Left            =   3600
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "maindoor3.frx":C042
         Caption         =   "Ãæá"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "maindoor3.frx":E1E9
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   2430
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "maindoor3.frx":10230
         Caption         =   "ÓÇÈÞ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "maindoor3.frx":1231B
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   1260
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "maindoor3.frx":14315
         Caption         =   "ÊÇáí"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "maindoor3.frx":16426
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "maindoor3.frx":18420
         Caption         =   "ÇÎíÑ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "maindoor3.frx":1A644
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
         TabIndex        =   5
         Top             =   450
         Width           =   4740
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   44
      Top             =   10395
      Visible         =   0   'False
      Width           =   15990
      _ExtentX        =   28205
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   90
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1058
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
      Picture         =   "maindoor3.frx":1C715
      Caption         =   "ÎÑæÌ"
      ButtonStyle     =   2
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "maindoor3.frx":1EAD3
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   600
      Left            =   1845
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   1058
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
      Picture         =   "maindoor3.frx":2142F
      Caption         =   "ÇáÛÇÁ ÇáÈíÇä"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "maindoor3.frx":23863
   End
   Begin Threed.SSCommand cmdType 
      Height          =   555
      Left            =   8460
      TabIndex        =   47
      Top             =   9135
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   979
      _Version        =   196610
      CaptionStyle    =   1
      ForeColor       =   -2147483641
      BackColor       =   16777215
      ActiveColors    =   -1  'True
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
      TabIndex        =   63
      Top             =   855
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "maindoorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, oSearchMember As New Search, oSearchRel As New Search
Dim CardTable As ADODB.Recordset, cFilter As String, cFile As String, cFile_Rel As String, oSearchType As New Search_empty
Const LoadMode = 1, DefineMode = 2
Private Sub cmdType_Click()
Set oSearchType = New Search_empty
TypeLookUp Me, oSearchType
End Sub
Private Sub CmdDel_Click()
'xCode.Caption = ""
'XCODE_REL.Caption = ""
'openCardTable
'myUndo
mydefine
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
mydefine
myload
End Sub
Private Sub CmdInform_Click()
If cmdType.Tag = 1 Then
    MemberLookupAll Me, oSearchMember
ElseIf cmdType.Tag = 2 Then
    Member_InLookupAll Me, oSearchMember
End If
End Sub

Private Sub cmdInformRel_Click()
If cmdType.Tag = 1 Then
    relLookupAll Me, oSearchRel
ElseIf cmdType.Tag = 2 Then
    relLookupAll_I Me, oSearchRel
End If
End Sub
Private Sub Form_Load()
SetKbLayout Lang_AR
openCon con
addIcons
'mydefine
cmdType.Tag = "1"
cmdType.Caption = "ÚÖæíÉ ÚÇãáÉ"
cFile = "file1_10"
cFile_Rel = "file1_11"
myUndo
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 * FROM " & cFile
If pCode <> "" Then cWhere = "CODE " & pSign & addvalue(pCode)

'cFilter = ""
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by CODE ASC"
End If
CardTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End Function
Private Sub myUndo()
On Error GoTo myerror
Dim cString As String, cWhere As String
If ValidNum(xCode.Caption) Then
    openCardTable xCode.Caption
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdNext_Click()
openCardTable xCode.Caption, ">"
If CardTable.EOF Then openCardTable xCode.Caption, "="
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.Caption, "<"
If CardTable.EOF Then openCardTable xCode.Caption, "="
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub myload()
Dim acode As Variant, loctable As New ADODB.Recordset, nIndex As Long, sPhoto As String, sPhotoRecord
Dim nUnPaid As Integer
Dim aPaid As Variant

ClearText
xCode.Caption = CardTable!CODE & ""
xdesca.Caption = CardTable!desca
If Not IsNull(CardTable!Type) Then
    xType_desca.Caption = GetField("SELECT DESCA FROM TYPE_CODES WHERE CODE = " & addvalue(CardTable!Type), con) & ""
Else
    xType_desca.Caption = ""
End If
xDate_Birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
Photo_main.Images.Clear
If cmdType.Tag = 1 Then
    If validPhoto(RetPhoto(xCode.Caption)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto(xCode.Caption))
    End If
    xLast_date.Caption = ""
        aPaid = Member_Paid(xCode.Caption, , con)
    If Not IsEmpty(aPaid) Then
        xUnPaid.Caption = unpaid_years(retFlag(aPaid, "year_code"), sSeason, con)
        xLast_date.Caption = myFormat_p(retFlag(aPaid, "date"))
        If mRound(xUnPaid.Caption) = 0 Then xUnPaid.Caption = "áÇ ÊæÌÏ ÓäæÇÊ ÛíÑ ãÓÏÏÉ"
        If retFlag(aPaid, "is_save") Then
            XLAST_PAID.Caption = "ÍÇÝÙ ÚÖæíÉ ÍÊí " & retFlag(aPaid, "year_desca") & ""
        Else
            XLAST_PAID.Caption = "ãÓÏÏ ÍÊí " & retFlag(aPaid, "year_desca") & ""
        End If
    Else
        XLAST_PAID.Caption = "áã íÓÏÏ ãä ÞÈá"
        xUnPaid.Caption = unpaid_years_count(xCode.Caption, sSeason, con)
    End If
ElseIf cmdType.Tag = 2 Then
    If validPhoto(RetPhoto_I(xCode.Caption)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto_I(xCode.Caption))
    End If
    cWhere = "INSTALL_BALANCE.CODE = " & xCode.Caption
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.Value - INSTALL_BALANCE.VALUE_PAID > 0"
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.DATE_DUE <= " & DateSq(Date)
    nUnPaid = mRound(GetField("SELECT Sum(INSTALL_BALANCE.INS_COUNT) AS Ins_Count FROM INSTALL_BALANCE  WHERE " & cWhere, con))
    xUnPaid_Install.Caption = IIf(nUnPaid = 0, "áÇ ÊæÌÏ ÇÞÓÇØ ãÊÃÎÑÉ", nUnPaid)
    xDate_Birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
    xLast_Date_I.Caption = myFormat_p(GetField("select dbo.f_last_year_date_install(" & xCode.Caption & ")", con))
    xFirst_Install.Caption = myFormat_p(GetField("select top 1 date_due from INSTALL_BALANCE WHERE VALUE - VALUE_PAID  > 0 AND CODE = " & xCode.Caption & " ORDER BY DATE_DUE ASC", con))
    If Not IsNull(CardTable!INSTALL_TYPE) Then
        xInstall_desca.Caption = GetField("SELECT DESCA FROM INSTALL_CODES WHERE CODE = " & addvalue(CardTable!INSTALL_TYPE), con) & ""
    Else
        xInstall_desca.Caption = ""
    End If
End If
Handlecontrols LoadMode
MyLoadPhotos
End Sub
Sub Handlecontrols(nMode)
aRecords = retRecords(xCode.Caption)
nRecord = Val(retFlag(aRecords, "record") & "")
nRecords = Val(retFlag(aRecords, "records") & "")
If nMode = LoadMode Then
    xRecord_No.Caption = ArbString("ÓÌá " & nRecord & " ãä " & nRecords)
Else
    xRecord_No.Caption = "áÇ ÊæÌÏ ÓÌáÇÊ"
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2
End Sub
Private Function retRecords(pCode) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If ValidNum(pCode) Then
    cString = "SELECT SUM(1) AS records,SUM(CASE WHEN CODE <= " & pCode & " THEN 1 ELSE 0 END) AS record"
Else
    cString = "SELECT SUM(1) AS records"
End If
cString = cString & " FROM " & cFile & turn(cFilter, " WHERE ") & cFilter
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    retRecords = AddFlag(Empty, "records", Val(loctable!records & ""))
    If ValidNum(pCode) Then retRecords = AddFlag(retRecords, "record", Val(loctable!Record & ""))
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set oSearchMember = Nothing
closeCon con
Set maindoorfrm = Nothing
Err.Clear
End Sub
Private Sub Photo_main_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Tag <> "" Then
    XCODE_REL.Caption = Source.Tag
    myUndo
ElseIf Source.Tag = -1 Then
    XCODE_REL.Caption = ""
    myUndo
End If
End Sub
Private Function MyLoadPhotos() As Boolean
Dim loctable As New ADODB.Recordset, cString As String, cWhere As String, i As Long, sPhotoRecord As String
For i = 1 To Photo1.UBound
    Photo1(i).Images.Clear
    Photo1(i).Tag = ""
    xdesca1(i).Caption = ""
    Photo1(i).Visible = False
    xdesca1(i).Visible = False
Next
cString = "SELECT " & cFile_Rel & ".* " & _
           " FROM " & cFile_Rel

cWhere = "MEMBER = " & xCode.Caption
'If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " order by CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
i = 0
Do Until loctable.EOF
    i = i + 1
    If i <= Photo1.UBound Then
        Photo1(i).Visible = True
        Photo1(i).Tag = loctable!CODE
        xdesca1(i).Visible = True
        xdesca1(i).Caption = loctable!desca & ""
        sPhotoRecord = loctable!MEMBER & "-" & loctable!CODE
        
        If cmdType.Tag = 1 Then
            If validPhoto(RetPhoto(sPhotoRecord)) Then
                Photo1(i).Import.FromFile RetPhoto(sPhotoRecord)
             End If
        ElseIf cmdType.Tag = 2 Then
            If validPhoto(RetPhoto_I(sPhotoRecord)) Then
                Photo1(i).Import.FromFile RetPhoto_I(sPhotoRecord)
             End If
        End If
    End If
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub mydefine()
Dim i As Long
For i = 1 To Photo1.UBound
    Photo1(i).Images.Clear
    Photo1(i).Tag = ""
    xdesca1(i).Caption = ""
    Photo1(i).Visible = False
    xdesca1(i).Visible = False
Next
Photo_main.Images.Clear
xCode.Caption = ""
ClearText
xRecord_No.Caption = "áÇ ÊæÌÏ ÈíÇäÇÊ"
Handlecontrols DefineMode
End Sub
Private Sub xBarCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If cmdGo.Enabled Then cmdGo_Click
End If
End Sub
Sub myProc()
If ActiveControl.Name = Me.cmdInform.Name Then
    xCode.Caption = oSearchMember.grid1.TextMatrix(oSearchMember.grid1.Row, 0)
    XCODE_REL.Caption = ""
    oSearchMember.Hide
ElseIf ActiveControl.Name = cmdInformRel.Name Then
    xCode.Caption = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 0)
    XCODE_REL.Caption = oSearchRel.grid1.TextMatrix(oSearchRel.grid1.Row, 1)
    oSearchRel.Hide
ElseIf ActiveControl.Name = cmdType.Name Then
    cmdType.Tag = oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 0)
    cmdType.Caption = IIf(cmdType.Tag = "", "äæÚ ÇáÚÖæíÉ", oSearchType.grid1.TextMatrix(oSearchType.grid1.Row, 1))
    If cmdType.Tag = 1 Then
        cFile = "FILE1_10"
        cFile_Rel = "file1_11"
        cFilter = ""
        FrameInstall.Visible = False
        FrameEmp.Visible = True
    ElseIf cmdType.Tag = 2 Then
        cFile = "FILE2_10"
        cFile_Rel = "file2_11"
        cFilter = "(COALESCE(FILE2_10.STATUS,0) <= 2)  "
        FrameInstall.Visible = True
        FrameEmp.Visible = False
    End If
    oSearchType.Hide
End If
myUndo
End Sub
Private Sub CloseData()
'closeCon con
End Sub
Private Sub TypeLookUp(oForm As Form, oSearch As Form, Optional bAddRow As Boolean = False)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "SELECT TYPE_CODES_SPORT.CODE,TYPE_CODES_SPORT.DESCA" & _
                  " From TYPE_CODES_SPORT "

Generalarray(2) = "Order by CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "ÇáäæÚ"
listarray(0, 1) = "(%%TYPE_CODES_SPORT.DESCA%%)"

GrdArray(0, 0) = "ÇáßæÏ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÇáäæÚ"
GrdArray(1, 1) = 5500

Dim aRow As Variant
If bAddRow Then
    aRow = AddFlag(Empty, "text", "äæÚ ÇáÚÖæíÉ")
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.sCaption = "ÅÓÊÚáÇã ÇäæÇÚ ÇáÚÖæíÉ"
oSearch.Show 1
End Sub
Private Sub ClearText()
Photo_main.Images.Clear
xLast_date.Caption = ""
XLAST_PAID.Caption = ""
xUnPaid.Caption = ""
xInstall_desca.Caption = ""
xUnPaid_Install.Caption = ""

xdesca.Caption = ""
xType_desca.Caption = ""
xLast_Date_I.Caption = ""
xFirst_Install.Caption = ""
xDate_Birth.Caption = ""
End Sub
Private Sub Photo1_Click(Index As Integer)
member_relfrm.sMember = xCode.Caption
member_relfrm.sCode = Photo1(Index).Tag
member_relfrm.nType = Val(cmdType.Tag)
member_relfrm.Show 1
End Sub
Private Sub Photo1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.MousePointer = 16
'SetCursor LoadCursor(0, IDC_HAND)
End Sub
Private Sub addIcons()
If Dir(App.Path & "\sys_img\hand.ico") = "" Then Exit Sub
Dim i As Long
For i = 1 To Me.Photo1.UBound
    Set Photo1(i).MouseIcon = LoadPicture(App.Path & "\sys_img\hand.ico")
Next
End Sub
