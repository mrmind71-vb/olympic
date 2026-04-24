VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8595
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
   ScaleHeight     =   4200
   ScaleWidth      =   8595
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
      Left            =   90
      Picture         =   "rpitem.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   3465
      Width           =   1500
   End
   Begin VB.CommandButton CMD_FIX 
      Caption         =   "÷»ÿ «· ﬂ·›…"
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
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3465
      Width           =   1590
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
      Height          =   3390
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   8430
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   2
         Left            =   4275
         TabIndex        =   1
         Top             =   630
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ—ﬂ… ’‰› ·› —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   3
         Left            =   4275
         TabIndex        =   2
         Top             =   1080
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ›’Ì·Ì  ÕÊÌ·«  „‰ „Œ“‰ «·Ì „Œ“‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   4
         Left            =   4275
         TabIndex        =   3
         Top             =   1530
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ì «· ÕÊÌ·«  „‰ „Œ“‰ «·Ì „Œ“‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   7
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ›’Ì·Ì «—»«Õ «·„»Ì⁄« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   1
         Left            =   4275
         TabIndex        =   0
         Top             =   180
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ﬁÌÌ„ «—’œ… «·‘—ﬂ…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   6
         Left            =   4275
         TabIndex        =   6
         Top             =   2430
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ã„«·Ì „‘ —Ì«  Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   5
         Left            =   4275
         TabIndex        =   7
         Top             =   1980
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ì  ÕÊÌ·«  ÌÊ„Ï"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   8
         Left            =   90
         TabIndex        =   8
         Top             =   675
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«—»«Õ „»Ì⁄«  ÌÊ„Ï"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   9
         Left            =   90
         TabIndex        =   9
         Top             =   1575
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ì «—»«Õ „»Ì⁄«  «·„⁄«—÷"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   11
         Left            =   45
         TabIndex        =   11
         Top             =   2610
         Visible         =   0   'False
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "›—Êﬁ «·Ã—œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   20
         Left            =   90
         TabIndex        =   12
         Top             =   1125
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈Ã„«·Ï „»Ì⁄«  ‘Â—Ï"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   12
         Left            =   4275
         TabIndex        =   14
         Top             =   2880
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«’‰«› —’ÌœÂ« «ﬁ· „‰ «Ê Ì”«ÊÌ Õœ «·ÿ·»"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   420
         Index           =   10
         Left            =   90
         TabIndex        =   15
         Top             =   2025
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   741
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ã„«·Ì ‰ﬁœÌ… ÌÊ„Ì… Œ·«· › —…"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New adodb.Connection
Dim nOption As Integer
Private Sub cmdApply_Click()
End Sub
Private Sub CMD_FIX_Click()
Dim oCost As New cost_fixfrm
oCost.Show 1
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub CmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
    rpitem1.Show 1
Case 2
    rpitem2.Show 1
Case 3
    rpItem3.Show 1
Case 4
    rpItem4.Show 1
Case 5
    rpItem5.Show 1
Case 6
    rpitem6.Show 1
Case 7
    rpsales1.Caption = cmdgo(Index).Caption
    rpsales1.Show 1
Case 8
    rpSales2.Caption = cmdgo(Index).Caption
    rpSales2.Show 1
Case 20
    rpSales20.Caption = cmdgo(Index).Caption
    rpSales20.Show 1
Case 9
    rpSales3.Caption = cmdgo(Index).Caption
    rpSales3.Show 1
Case 10
    rpSales4.Caption = cmdgo(Index).Caption
    rpSales4.Show 1
Case 11
    rpitem11.Caption = cmdgo(Index).Caption
    rpitem11.Show 1
Case 12
    rpitem12.Caption = cmdgo(Index).Caption
    rpitem12.Show 1
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
Private Sub Form_Load()
    openCon con
    cmdgo(7).Visible = bopt1
    cmdgo(8).Visible = bopt1
    cmdgo(9).Visible = bopt1
    CMD_FIX.Visible = bopt1
    
End Sub
Private Sub fix1()
    cCaption = Me.Caption
    openCon con
    Dim loctable As New adodb.Recordset
    cString = "Select FILE1_10.ITEM from file1_10"
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    
    nCount = loctable.RecordCount
    Dim i As Long
    con.BeginTrans
    Do Until loctable.EOF
        i = i + 1
        Me.Caption = cCaption & i & "  „‰ " & nCount
        nCost = LastCost(loctable!Item, con)
        If nCost <> 0 Then
            con.Execute " UPDATE FILE1_10 SET FILE1_10.COST = " & nCost & _
                        " WHERE FILE1_10.item = " & MyParn(loctable!Item)
        End If
        loctable.MoveNext
    Loop
    con.CommitTrans
    Me.Caption = cCaption
    MsgBox " „ ÷»ÿ «· ﬂ·›… »‰Ã«Ã"
lastsub:
    closeCon con
    Exit Sub
myerror:
Me.Caption = cCaption
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
    GoTo lastsub
End Sub
Private Sub fix2()
    cCaption = Me.Caption
    openCon con
    Dim loctable As New adodb.Recordset
    cString = "Select FILE6_20.ITEM,FILE6_20H.DATE,FILE6_20.ID FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    
    nCount = loctable.RecordCount
    Dim i As Long
    con.BeginTrans
    Do Until loctable.EOF
        i = i + 1
        Me.Caption = cCaption & i & "  „‰ " & nCount
        nCost = LastCostDate(loctable!Item, Format(loctable!Date, "yyyy-mm-dd"), con)
        If nCost <> 0 Then
            con.Execute " UPDATE FILE6_20 SET FILE6_20.COST = " & nCost & _
                        " WHERE FILE6_20.ID = " & loctable!ID
        End If
        loctable.MoveNext
    Loop
    con.CommitTrans
    Me.Caption = cCaption
    MsgBox " „ ÷»ÿ «· ﬂ·›… »‰Ã«Ã"
lastsub:
    closeCon con
    Exit Sub
myerror:
Me.Caption = cCaption
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
    GoTo lastsub
End Sub
