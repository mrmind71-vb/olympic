VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form sportfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÇÓÊÚáÇã ÇÚÖÇÁ ÇáäÇÏí"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15885
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "sport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   43
      Top             =   8190
      Width           =   4785
      Begin Threed.SSCommand cmdFirst 
         Default         =   -1  'True
         Height          =   420
         Left            =   3600
         TabIndex        =   44
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
         Picture         =   "sport.frx":12632
         Caption         =   "Ãæá"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "sport.frx":147D9
      End
      Begin Threed.SSCommand cmdPrevious 
         Height          =   420
         Left            =   2430
         TabIndex        =   45
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
         Picture         =   "sport.frx":16820
         Caption         =   "ÓÇÈÞ"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "sport.frx":1890B
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   420
         Left            =   1260
         TabIndex        =   46
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
         Picture         =   "sport.frx":1A905
         Caption         =   "ÊÇáí"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "sport.frx":1CA16
      End
      Begin Threed.SSCommand cmdLast 
         Height          =   420
         Left            =   0
         TabIndex        =   47
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
         Picture         =   "sport.frx":1EA10
         Caption         =   "ÇÎíÑ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   1
         PictureDisabledFrames=   1
         PictureDisabled =   "sport.frx":20C34
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
         TabIndex        =   48
         Top             =   450
         Width           =   4740
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
      TabIndex        =   12
      Top             =   90
      Width           =   8340
      Begin ImgXCtrl6.ImgXCtrl Photo1 
         DragIcon        =   "sport.frx":22D05
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   1
         Left            =   6165
         TabIndex        =   13
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
         DragIcon        =   "sport.frx":23147
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   2
         Left            =   4185
         TabIndex        =   14
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
         DragIcon        =   "sport.frx":23589
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   3
         Left            =   2205
         TabIndex        =   15
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
         DragIcon        =   "sport.frx":239CB
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   4
         Left            =   225
         TabIndex        =   16
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
         DragIcon        =   "sport.frx":23E0D
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   5
         Left            =   6165
         TabIndex        =   17
         Tag             =   "-1"
         Top             =   3240
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
         DragIcon        =   "sport.frx":2424F
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   6
         Left            =   4185
         TabIndex        =   18
         Tag             =   "-1"
         Top             =   3240
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
         DragIcon        =   "sport.frx":24691
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   7
         Left            =   2205
         TabIndex        =   19
         Tag             =   "-1"
         Top             =   3240
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
         DragIcon        =   "sport.frx":24AD3
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   8
         Left            =   225
         TabIndex        =   20
         Tag             =   "-1"
         Top             =   3240
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
         DragIcon        =   "sport.frx":24F15
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   9
         Left            =   6165
         TabIndex        =   21
         Tag             =   "-1"
         Top             =   6525
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
         DragIcon        =   "sport.frx":25357
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   10
         Left            =   4185
         TabIndex        =   22
         Tag             =   "-1"
         Top             =   6525
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
         DragIcon        =   "sport.frx":25799
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   11
         Left            =   2205
         TabIndex        =   23
         Tag             =   "-1"
         Top             =   6525
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
         DragIcon        =   "sport.frx":25BDB
         DragMode        =   1  'Automatic
         Height          =   2310
         Index           =   12
         Left            =   225
         TabIndex        =   24
         Tag             =   "-1"
         Top             =   6525
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
         Height          =   600
         Index           =   12
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         Index           =   9
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   33
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
         Index           =   8
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   32
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
         Index           =   7
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   31
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
         TabIndex        =   30
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
         Index           =   5
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   29
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
         Index           =   4
         Left            =   225
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
         Height          =   600
         Index           =   2
         Left            =   4185
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
         Height          =   600
         Index           =   1
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2475
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   5565
      ScaleWidth      =   4845
      TabIndex        =   3
      Top             =   495
      Width           =   4875
      Begin ImgXCtrl6.ImgXCtrl Photo_main 
         DragIcon        =   "sport.frx":2601D
         Height          =   3660
         Left            =   675
         TabIndex        =   4
         Tag             =   "-1"
         Top             =   135
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   6456
         BackColor       =   16777215
         BorderStyle     =   3
         AutoZoom        =   -1  'True
         LicenseUserName =   "mrmind71"
         LicenseRegCode  =   "’§Ò½»º­½³«±Òª¼¯«´¾®¯UBOR-FEOEONZI-EPCP6gI"
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáãíáÇÏ"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   5130
         Width           =   1005
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   5085
         Width           =   3435
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3870
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÑÞã "
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3915
         Width           =   735
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3870
         Width           =   1995
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   4275
         Width           =   3435
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÇÓã"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   4275
         Width           =   690
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   4680
         Width           =   3435
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇáÞÑÇÈÉ"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   4680
         Width           =   645
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
      TabIndex        =   0
      Top             =   7515
      Width           =   4830
      Begin Threed.SSCommand cmdInform 
         Height          =   510
         Left            =   2430
         TabIndex        =   1
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
         Picture         =   "sport.frx":2645F
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÇÓÇÓí"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sport.frx":2882A
      End
      Begin Threed.SSCommand cmdInformRel 
         Height          =   510
         Left            =   45
         TabIndex        =   2
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
         Picture         =   "sport.frx":2A8D3
         Caption         =   "ÈÍË ÈÇáÚÖæ ÇáÊÇÈÚ"
         ButtonStyle     =   3
         PictureAlignment=   9
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sport.frx":2CC9E
      End
   End
   Begin ComctlLib.ProgressBar Prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   37
      Top             =   10470
      Visible         =   0   'False
      Width           =   15885
      _ExtentX        =   28019
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   90
      TabIndex        =   40
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
      Picture         =   "sport.frx":2ED47
      Caption         =   "ÎÑæÌ"
      ButtonStyle     =   2
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "sport.frx":31105
   End
   Begin Threed.SSCommand cmdDel 
      Height          =   600
      Left            =   1845
      TabIndex        =   41
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
      Picture         =   "sport.frx":33A61
      Caption         =   "ÇáÛÇÁ ÇáÈíÇä"
      ButtonStyle     =   3
      PictureAlignment=   9
      BevelWidth      =   0
      PictureDisabledFrames=   1
      ShapeSize       =   1
      PictureDisabled =   "sport.frx":35E95
   End
   Begin Threed.SSCommand cmdType 
      Height          =   555
      Left            =   8460
      TabIndex        =   42
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
   Begin VB.PictureBox FrameInstall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   8505
      RightToLeft     =   -1  'True
      ScaleHeight     =   1335
      ScaleWidth      =   4800
      TabIndex        =   49
      Top             =   6120
      Visible         =   0   'False
      Width           =   4830
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
         TabIndex        =   65
         Top             =   90
         Width           =   1500
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
         TabIndex        =   64
         Top             =   900
         Width           =   3120
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáÞÓØ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÚÏÏ ÇÞÓÇØ ãÊÃÎÑÉ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   540
         Width           =   1410
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
         TabIndex        =   52
         Top             =   495
         Width           =   3120
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇÎÑ ÓÏÇÏ"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   135
         Width           =   1185
      End
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
         TabIndex        =   50
         Top             =   90
         Width           =   1590
      End
   End
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
         TabIndex        =   62
         Top             =   90
         Width           =   3570
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇÎÑ ãæÓã"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
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
         TabIndex        =   60
         Top             =   495
         Width           =   3570
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÊÇÑíÎ ÇáÓÏÇÏ"
         Height          =   330
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   540
         Width           =   1005
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
         TabIndex        =   58
         Top             =   900
         Width           =   3120
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÓäæÇÊ ÛíÑ ãÓÏÏÉ"
         Height          =   330
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   945
         Width           =   1365
      End
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
      TabIndex        =   39
      Top             =   450
      Visible         =   0   'False
      Width           =   2040
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
      TabIndex        =   38
      Top             =   855
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "sportfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, oSearchMember As New Search, oSearchRel As New Search
Dim CardTable As ADODB.Recordset, cFilter As String, cFile As String, oSearchType As New Search_empty
Const LoadMode = 1, DefineMode = 2
Private Sub cmdType_Click()
Set oSearchType = New Search_empty
TypeLookUp Me, oSearchType
End Sub
Private Sub CmdDel_Click()
xCode.Caption = ""
XCODE_REL.Caption = ""
openCardtable
myUndo
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
myDefine
myload
End Sub
Private Sub CmdInform_Click()
If cmdType.Tag = 1 Then
    MemberLookupAll Me, oSearchMember
ElseIf cmdType.Tag = 2 Then
    Member_InLookupAll Me, oSearchMember
End If
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub

Private Sub cmdInformRel_Click()
If cmdType.Tag = 1 Then
    relLookupAll Me, oSearchRel
ElseIf cmdType.Tag = 2 Then
    relLookupAll_I Me, oSearchRel
End If
End Sub

Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub Form_Load()
SetKbLayout Lang_AR

sCatalog = "OLYMPIC"
sMdfName = "OLYMPIC"
strCon = LoadConString
openCon con

aSeason = GetFields("Select Top 1 * from years_codes where Date1 <= " & DateSq(Date) & " and date2 >= " & DateSq(Date) & " order by date1 desc", con)
If IsEmpty(aSeason) Then
    aSeason = GetFields("Select Top 1 * from years_codes order by date1 desc", con)
End If
sSeason = retFlag(aSeason, "code")

sDate_Season = myFormat(retFlag(aSeason, "date1"))


myDefine
cmdType.Tag = "1"
cmdType.Caption = "ÚÖæíÉ ÚÇãáÉ"
cFile = "Members"
End Sub
Private Function openCardtable()
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cString = "SELECT  * " & _
           " FROM " & cFile

cFilter = ""
cFilter = "[MEMBER] = " & addvalue(xCode.Caption)
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

MyLoadPhotos
End Function
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    myDefine
Else
    If XCODE_REL.Caption <> "" Then
        CardTable.Find "CODE = " & addvalue(XCODE_REL.Caption), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveFirst
    Else
        CardTable.MoveFirst
    End If
    myload
End If
End Sub
Private Sub myload()
Dim acode As Variant, loctable As New ADODB.Recordset, nIndex As Long, sPhoto As String, sPhotoRecord
Dim nUnPaid As Integer
Dim aPaid As Variant

ClearText

xMember.Caption = CardTable!MEMBER & ""
xMember.Tag = CardTable!MEMBER & ""
If CardTable!flag = 2 Then
    xMember.Caption = xMember.Caption & "-" & CardTable!CODE
    xMember.Tag = xMember.Tag
    XCODE_REL.Caption = CardTable!CODE
Else
    XCODE_REL.Caption = ""
End If
xDesca.Caption = CardTable!Desca & ""
xType_desca.Caption = CardTable!TYPE_desca & ""
xDate_birth.Caption = myFormat_p(CardTable!DATE_BIRTH)

xRelation_Desca.Caption = CardTable!RELATION_DESCA & ""
xRecord_No.Caption = "ÓÌá " & CardTable.AbsolutePosition & " ãä " & CardTable.RecordCount

Photo_main.Images.Clear
If cmdType.Tag = 1 Then
    xLast_date.Caption = ""
    
    
    If validPhoto(RetPhoto(xMember.Tag)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto(xMember.Caption))
    
    End If
    
    aPaid = Member_Paid(xMember.Tag, , con)
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
        xUnPaid.Caption = unpaid_years_count(xMember.Tag, sSeason, con)
    End If
ElseIf cmdType.Tag = 2 Then
    cWhere = "INSTALL_BALANCE.CODE = " & xMember.Tag
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.Value - INSTALL_BALANCE.VALUE_PAID > 0"
    cWhere = cWhere & " AND " & "INSTALL_BALANCE.DATE_DUE <= " & DateSq(Date)
    nUnPaid = mRound(GetField("SELECT Sum(INSTALL_BALANCE.INS_COUNT) AS Ins_Count FROM INSTALL_BALANCE  WHERE " & cWhere, con))
    xUnPaid_Install.Caption = IIf(nUnPaid = 0, "áÇ ÊæÌÏ ÇÞÓÇØ ãÊÃÎÑÉ", nUnPaid)
    xDate_birth.Caption = myFormat_p(CardTable!DATE_BIRTH)
    xLast_Date_I.Caption = myFormat_p(GetField("select dbo.f_last_year_date_install(" & xMember.Tag & ")", con))
    xFirst_Install.Caption = myFormat_p(GetField("select top 1 date_due from INSTALL_BALANCE WHERE VALUE - VALUE_PAID  > 0 AND CODE = " & xMember.Tag & " ORDER BY DATE_DUE ASC", con))
    xInstall_desca.Caption = CardTable!install_desca & ""
    xRelation_Desca.Caption = CardTable!RELATION_DESCA & ""
    If validPhoto(RetPhoto_I(xMember.Tag)) Then
        Set Photo_main.Picture = LoadPicture(RetPhoto_I(xMember.Tag))
    End If
End If
Handlecontrols LoadMode
End Sub
Sub Handlecontrols(nMode)
If nMode = DefineMode Then
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdFirst.Enabled = False
Else
    cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
    cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
    cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
    cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
End If
End Sub
Function aUnMyCodeBar(sCode) As Variant
Dim nVal1 As Integer, nVal2 As Integer, nValSub1 As Integer, nValSub2 As Integer
Dim nNumber1 As String, nNumber2 As String, nNumber3 As String


If Trim(sCode) = "" Then Exit Function
Dim aRet As Variant, aSplit As Variant
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
aRet = AddFlag(Empty, "MEMBER", aSplit(0))
If UBound(aSplit) = 1 Then aRet = AddFlag(aRet, "CODE", aSplit(1))
'aret = AddFlag(aret, "TYPE", IIf(Val(nNumber3) = 0, "", Val(nNumber3)))
aUnMyCodeBar = aRet
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set oSearchMember = Nothing
'addSetting xDrive.Name, xDrive.Text, TempSave(Me)
closeCon con
Set maindoorfrm = Nothing
Err.Clear
End
End Sub

Private Sub Label9_Click()

End Sub

Private Sub Label11_Click()

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
cString = "SELECT " & cFile & ".* " & _
           " FROM " & cFile
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " order by CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
i = 0
Do Until loctable.EOF
    i = i + 1
    If i <= Photo1.UBound Then
        Photo1(i).Visible = True
        xdesca1(i).Visible = True
        xdesca1(i).Caption = loctable!Desca & ""
        sPhotoRecord = loctable!MEMBER
        If loctable!flag <> 1 Then
            sPhotoRecord = sPhotoRecord & "-" & loctable!CODE
            Photo1(i).Tag = loctable!CODE
        Else
            Photo1(i).Tag = "-1"
        End If
        If cmdType.Tag = 1 Then
            If validPhoto(RetPhoto(sPhotoRecord)) Then
                Photo1(i).Import.FromFile RetPhoto(sPhotoRecord)
             End If
        ElseIf cmdType.Tag = 2 Then
            If validPhoto(RetPhoto_I(sPhotoRecord)) Then
                Photo1(i).Import.FromFile RetPhoto_I(sPhotoRecord)
             End If
        ElseIf cmdType.Tag = 3 Then
            If validPhoto(RetPhotoh(sPhotoRecord)) Then
                Photo1(i).Import.FromFile RetPhotoh(sPhotoRecord)
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
Private Sub myDefine()
Dim i As Long
For i = 1 To Photo1.UBound
    Photo1(i).Images.Clear
    Photo1(i).Tag = ""
    xdesca1(i).Caption = ""
    Photo1(i).Visible = False
    xdesca1(i).Visible = False
Next
Photo_main.Images.Clear
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
        cFile = "MEMBERS"
        FrameInstall.Visible = False
        FrameEmp.Visible = True
    ElseIf cmdType.Tag = 2 Then
        cFile = "MEMBERS_INV"
        FrameInstall.Visible = True
        FrameEmp.Visible = False
    End If
    'cmdInform.Enabled = cmdType.Tag <> ""
    'cmdInformRel.Enabled = cmdType.Tag <> ""
    oSearchType.Hide
End If
openCardtable
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
xMember.Caption = ""
xMember.Tag = ""
xCode.Caption = ""
xRelation_Desca.Caption = ""
XCODE_REL.Caption = ""
xLast_date.Caption = ""
xUnPaid.Caption = ""
xInstall_desca.Caption = ""
xUnPaid_Install.Caption = ""
XLAST_PAID.Caption = ""

xDesca.Caption = ""
xType_desca.Caption = ""
xLast_Date_I.Caption = ""
xFirst_Install.Caption = ""
xDate_birth.Caption = ""
End Sub




