VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form reportfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— «·«⁄÷«¡ «·⁄«„·Ì‰"
   ClientHeight    =   6930
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   12840
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "report.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   6345
      UseMaskColor    =   -1  'True
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   12705
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   1
         Left            =   8955
         TabIndex        =   0
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "«⁄œ«œ «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   2
         Left            =   8955
         TabIndex        =   3
         Top             =   630
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "»Ì«‰«  «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   18
         Left            =   4770
         TabIndex        =   4
         Top             =   1980
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«»‰«¡  ⁄œÊ« ”‰ „⁄Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   15
         Left            =   4770
         TabIndex        =   6
         Top             =   630
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "»Ì«‰«   «· «»⁄Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   3
         Left            =   8955
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ €Ì— „ﬂ „·Ì »Ì«‰«  «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   4
         Left            =   8955
         TabIndex        =   8
         Top             =   1530
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ Õ”» «·ÊŸÌ›… -«·‘—ﬂ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   6
         Left            =   8955
         TabIndex        =   9
         Top             =   2430
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·«⁄÷«¡ Õ”» «·œÌ«‰…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   5
         Left            =   8955
         TabIndex        =   5
         Top             =   1980
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·«⁄÷«¡ Õ”» «·„ƒÂ·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   7
         Left            =   8955
         TabIndex        =   10
         Top             =   2880
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·«⁄÷«¡ Õ”» „Õ· «·«ﬁ«„…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   14
         Left            =   4770
         TabIndex        =   11
         Top             =   180
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«“Ê«Ã «·«⁄÷«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   16
         Left            =   4770
         TabIndex        =   12
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   " «»⁄Ì‰ €Ì— „ﬂ „·Ì »Ì«‰«  «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   17
         Left            =   4770
         TabIndex        =   13
         Top             =   1530
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄Ì«œ «·„Ì·«œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   11
         Left            =   8955
         TabIndex        =   14
         Top             =   4680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ «Œ— ”œ«œ ·Â„ ›Ï „Ê”„ „⁄Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   10
         Left            =   8955
         TabIndex        =   15
         Top             =   4230
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ ”œÊœ« „Ê”„ „⁄Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   12
         Left            =   8955
         TabIndex        =   16
         Top             =   5130
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ ·„ Ì”œœÊ« „Ê”„ „⁄Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   9
         Left            =   8955
         TabIndex        =   17
         Top             =   3780
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ Õ«›ŸÌ «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   8
         Left            =   8955
         TabIndex        =   18
         Top             =   3330
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·«⁄÷«¡ Õ”»  «—ÌŒ «·«· Õ«ﬁ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   19
         Left            =   4770
         TabIndex        =   19
         Top             =   3330
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·Ã„⁄Ì… «·⁄„Ê„Ì…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   13
         Left            =   8955
         TabIndex        =   20
         Top             =   5580
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«·«⁄÷«¡ Õ”» ›∆… «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   22
         Left            =   4770
         TabIndex        =   21
         Top             =   3780
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "≈Ã„«·Ì ”œ«œ „Ã„Ê⁄«  «·»‰Êœ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   23
         Left            =   4770
         TabIndex        =   22
         Top             =   4230
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   " ≈Ã„«·Ì «Ì’«·«  ”œ«œ «·«⁄÷«¡  ›’Ì·Ì"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   20
         Left            =   4770
         TabIndex        =   23
         Top             =   4680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "Œÿ«»«  «Œÿ«—"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   24
         Left            =   4770
         TabIndex        =   24
         Top             =   2430
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ ›«’·Ì «·⁄÷ÊÌ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   25
         Left            =   4770
         TabIndex        =   25
         Top             =   2880
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "«⁄÷«¡ Ãœœ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   21
         Left            =   4770
         TabIndex        =   26
         Top             =   5130
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "Œÿ«»«  «”ﬁ«ÿ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   31
         Left            =   4770
         TabIndex        =   27
         Top             =   5580
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "√⁄œ«œ «·«⁄÷«¡ –ﬂÊ— Ê«‰«À"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   32
         Left            =   180
         TabIndex        =   28
         Top             =   135
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   " ›’Ì·Ì „” Õﬁ«  ÷—Ì»… ﬁÌ„… „÷«›… "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   33
         Left            =   180
         TabIndex        =   29
         Top             =   585
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "≈Ã„«·Ì „” Õﬁ«  ÷—Ì»… „÷«›…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   34
         Left            =   180
         TabIndex        =   30
         Top             =   1035
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   " ·Ì›Ê‰«  «·«⁄÷«¡ «·„ Œ·›Ì‰ ⁄‰  ÷—Ì»… «·«ﬁ”«ÿ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   35
         Left            =   180
         TabIndex        =   31
         Top             =   1485
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "⁄‰Ê«Ì‰ «·«⁄÷«¡ «·„ Œ·›Ì‰ ⁄‰  ÷—Ì»… «·«ﬁ”«ÿ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   36
         Left            =   180
         TabIndex        =   32
         Top             =   1935
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   " «⁄œ«œ «·«⁄÷«¡ «·„ Œ·›Ì‰ Õ”» «·ﬁÌ„…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   37
         Left            =   180
         TabIndex        =   33
         Top             =   2385
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "”œ«œ ›—Êﬁ ﬁÌ„… „÷«›…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   38
         Left            =   180
         TabIndex        =   34
         Top             =   2790
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "”œ«œ ›—Êﬁ ﬁÌ„… „÷«›… ··«ﬁ”«ÿ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   39
         Left            =   180
         TabIndex        =   35
         Top             =   3240
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   " ›’Ì·Ì ›—Êﬁ ﬁÌ„… „÷«›… €Ì— „”Ã·…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   41
         Left            =   180
         TabIndex        =   36
         Top             =   3690
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "—’Ìœ ›—Êﬁ ﬁÌ„… „÷«›… ··«⁄÷«¡ «·⁄«„·Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   40
         Left            =   180
         TabIndex        =   37
         Top             =   4140
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "—’Ìœ ›—Êﬁ ﬁÌ„… „÷«›… ··«⁄÷«¡ «·„ﬁ”ÿÌ‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   42
         Left            =   180
         TabIndex        =   38
         Top             =   4590
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         BackColor       =   14737632
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
         Caption         =   "”œ«œ ›—Êﬁ ﬁÌ„… „÷«›… ··„ﬁ”ÿÌ‰ ‘Â—Ì"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   26
         Left            =   180
         TabIndex        =   39
         Top             =   5040
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "»Ì«‰«  «·«⁄÷«¡ »«·—ﬁ„ «·ﬁÊ„Ì Ê«·„Õ„Ê·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   30
         Left            =   180
         TabIndex        =   40
         Top             =   5490
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   741
         _Version        =   196610
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
         Caption         =   "”Ã· «·„—«Ã⁄…"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "reportfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nOption As Integer
Private Sub cmdApply_Click()
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
     reportfrm1.Show 1
Case 2
     reportfrm2.Show 1
Case 3
     reportfrm3.Show 1
Case 4
     reportfrm4.Show 1
Case 5
     reportfrm5.Show 1
Case 6
     reportfrm6.Show 1
Case 7
     reportfrm7.Show 1
Case 8
     reportfrm8.Show 1
Case 9
     reportfrm9.Show 1
Case 10
     reportfrm10.Show 1
Case 11
     reportfrm11.Show 1
Case 12
     reportfrm12.Show 1
Case 13
     reportfrm13.Show 1
Case 14
     reportfrm14.Show 1
Case 15
     reportfrm15.Show 1
Case 16
     reportfrm16.Show 1
Case 17
     reportfrm17.Show 1
Case 18
     reportfrm18.Show 1
Case 19
     reportfrm19.Show 1
Case 20
     reportfrm20.Show
Case 21
     reportfrm21.Show
Case 22
    grdpaid2.Show
Case 23
    grdpaid1.Show
Case 24
     reportfrm24.Show 1
Case 25
     reportfrm25.Show 1
Case 31
     reportfrm31.Show 1
Case 32
    grdTax1.Show
Case 33
    grdTax2.Show
Case 34
    grdTax3.Show
Case 35
    grdTax4.Show
Case 36
    grdTax5.Show
Case 37
    grdTax6.Show
Case 38
    grdTax7.Show
Case 39
    grdTax8.Show
Case 40
    grdTax9.Show
Case 41
    grdTax10.Show
Case 42
    grdTax20.Show
Case 30
     reportfrm30.Show 1
End Select
End Sub
Private Sub cmdgo_MouseEnter(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub

