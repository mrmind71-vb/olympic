VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form StorGrd 
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2895
      TabIndex        =   1
      Top             =   2010
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1890
      TabIndex        =   0
      Top             =   2010
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid StoreGrd 
      Height          =   1980
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   15
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   3493
      _Version        =   393216
      Cols            =   4
      ScrollTrack     =   -1  'True
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Left            =   135
      TabIndex        =   4
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label StoreLbl 
      Caption         =   "Label1"
      Height          =   195
      Left            =   3960
      TabIndex        =   3
      Top             =   2100
      Width           =   615
   End
End
Attribute VB_Name = "StorGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub StoreGrd_EnterCell()
Label1.Caption = StoreGrd.row
End Sub

