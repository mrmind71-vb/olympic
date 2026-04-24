VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form DaySalClosefrm 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3930
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì „ „»Ì⁄«  :"
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Ã„«·Ì „»Ì⁄«   :"
         Height          =   285
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·»«∆⁄"
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   3885
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
   End
End
Attribute VB_Name = "DaySalClosefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
