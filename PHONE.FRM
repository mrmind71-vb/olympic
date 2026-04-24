VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PHONE 
   Caption         =   "œ·Ì· «· ·Ì›Ê‰« "
   ClientHeight    =   6300
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   7695
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdPhone 
      BackColor       =   &H00E0E0E0&
      Caption         =   "»ÕÀ —ﬁ„"
      Height          =   390
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton CMD_PRINT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÿ»«⁄…"
      Height          =   390
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton CmdLast 
      BackColor       =   &H00C0FFFF&
      Caption         =   "√ŒÌ—"
      Height          =   390
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5700
      Width           =   1065
   End
   Begin VB.CommandButton CmdFirst 
      BackColor       =   &H00C0FFFF&
      Caption         =   "√Ê·"
      Height          =   390
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5700
      Width           =   1140
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "·«Õﬁ"
      Height          =   390
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00C0FFFF&
      Caption         =   "”«»ﬁ"
      Height          =   390
      Left            =   3975
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5700
      Width           =   1290
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdInform 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«” ⁄·«„"
      Height          =   390
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E0E0E0&
      Caption         =   " —«Ã⁄"
      Height          =   390
      Left            =   2790
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«÷«›…"
      Height          =   390
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton CmdDel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ–›"
      Height          =   390
      Left            =   930
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Õ›Ÿ"
      Height          =   390
      Left            =   1860
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.TextBox xNameA 
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
      Height          =   315
      Left            =   510
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1062
      Width           =   5988
   End
   Begin VB.TextBox xCOMP 
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
      Height          =   315
      Left            =   510
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   5988
   End
   Begin VB.TextBox xAddr 
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
      Height          =   315
      Left            =   510
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1464
      Width           =   5988
   End
   Begin VB.TextBox xJob 
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
      Height          =   315
      Left            =   510
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1866
      Width           =   5988
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5235
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1260
   End
   Begin VB.TextBox xPHONE6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   510
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   1530
   End
   Begin VB.TextBox xDesc5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2100
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3240
      Width           =   1800
   End
   Begin VB.TextBox xMemo 
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
      Height          =   1032
      Left            =   525
      MaxLength       =   250
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4050
      Width           =   5988
   End
   Begin VB.TextBox xDesc3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   5805
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3600
      Width           =   1800
   End
   Begin VB.TextBox xPHONE5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   510
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3240
      Width           =   1530
   End
   Begin VB.TextBox xDesc2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   5805
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   1800
   End
   Begin VB.TextBox xPhone1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4155
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   1530
   End
   Begin VB.TextBox xPhone2 
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
      Height          =   300
      Left            =   4155
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3240
      Width           =   1530
   End
   Begin VB.TextBox xPHONE4 
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
      Height          =   300
      Left            =   510
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2880
      Width           =   1530
   End
   Begin VB.TextBox xDesc1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   5805
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   1800
   End
   Begin VB.TextBox xDesc4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2100
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   1800
   End
   Begin VB.TextBox xPhone3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4155
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3600
      Width           =   1530
   End
   Begin VB.TextBox xDesc6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2100
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·‘—ﬂ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6690
      TabIndex        =   22
      Top             =   2295
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊŸÌ›…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6690
      TabIndex        =   21
      Top             =   1905
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄‰Ê«‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6690
      TabIndex        =   20
      Top             =   1515
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·√”„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6690
      TabIndex        =   19
      Top             =   1110
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬂÊœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   6690
      TabIndex        =   18
      Top             =   720
      Width           =   435
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   6690
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   0
      X2              =   8775
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   0
      X2              =   8775
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   4050
      X2              =   4050
      Y1              =   4095
      Y2              =   2775
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      X1              =   450
      X2              =   7800
      Y1              =   2700
      Y2              =   2700
   End
End
Attribute VB_Name = "PHONE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cardtable As Recordset, TempTable As Recordset
Dim formMode As Byte, lValidNow As Boolean
Const Editmode = 1, addmode = 2, EmptyMode = 3
Dim ValidString As String
Sub myDefine()
xNameA.Text = ""
xAddr.Text = ""
xJob.Text = ""
xCOMP.Text = ""
xPhone1.Text = ""
xPhone2.Text = ""
xPhone3.Text = ""
xPHONE4.Text = ""
xPHONE5.Text = ""
xPHONE6.Text = ""
xDesc1.Text = "„‰“·"
xDesc2.Text = "⁄„·"
xDesc3.Text = "„Ê»Ì·"
xDesc4.Text = ""
xDesc5.Text = ""
xDesc6.Text = ""
xMemo.Text = ""
End Sub
Sub myProc()
    cardtable.FindFirst "Code = " & MyParn(GrdText(Search.Grid1, 0))
    MyLoad
End Sub
Sub MyLoad()
xCode.Text = TurnValue(cardtable.CODE, Null, "")
xNameA.Text = TurnValue(cardtable.NAMEA, Null, "")
xJob.Text = TurnValue(cardtable.JOB, Null, "")
xAddr.Text = TurnValue(cardtable.Addr, Null, "")
xCOMP.Text = TurnValue(cardtable.COMP, Null, "")
xPhone1.Text = TurnValue(cardtable.PHONE1, Null, "")
xPhone2.Text = TurnValue(cardtable.PHONE2, Null, "")
xPhone3.Text = TurnValue(cardtable.PHONE3, Null, "")
xPHONE4.Text = TurnValue(cardtable.PHONE4, Null, "")
xPHONE5.Text = TurnValue(cardtable.PHONE5, Null, "")
xPHONE6.Text = TurnValue(cardtable.PHONE6, Null, "")
xDesc1.Text = TurnValue(cardtable.DESC1, Null, "")
xDesc2.Text = TurnValue(cardtable.DESC2, Null, "")
xDesc3.Text = TurnValue(cardtable.DESC3, Null, "")
xDesc4.Text = TurnValue(cardtable.DESC4, Null, "")
xDesc5.Text = TurnValue(cardtable.DESC5, Null, "")
xDesc6.Text = TurnValue(cardtable.DESC6, Null, "")
xMemo.Text = TurnValue(cardtable.Memo, Null, "")
End Sub
Sub MyReplace()
cardtable.CODE = TurnValue(xCode.Text, "", Null)
cardtable.NAMEA = TurnValue(xNameA.Text, "", Null)
cardtable.COMP = TurnValue(xCOMP.Text, "", Null)
cardtable.Addr = TurnValue(xAddr.Text, "", Null)
cardtable.JOB = TurnValue(xJob.Text, "", Null)
cardtable.PHONE1 = TurnValue(xPhone1.Text, "", Null)
cardtable.PHONE2 = TurnValue(xPhone2.Text, "", Null)
cardtable.PHONE3 = TurnValue(xPhone3.Text, "", Null)
cardtable.PHONE4 = TurnValue(xPHONE4.Text, "", Null)
cardtable.PHONE5 = TurnValue(xPHONE5.Text, "", Null)
cardtable.PHONE6 = TurnValue(xPHONE6.Text, "", Null)
cardtable.DESC1 = TurnValue(xDesc1.Text, "", Null)
cardtable.DESC2 = TurnValue(xDesc2.Text, "", Null)
cardtable.DESC3 = TurnValue(xDesc3.Text, "", Null)
cardtable.DESC4 = TurnValue(xDesc4.Text, "", Null)
cardtable.DESC5 = TurnValue(xDesc5.Text, "", Null)
cardtable.DESC6 = TurnValue(xDesc6.Text, "", Null)
cardtable.Memo = TurnValue(xMemo.Text, "", Null)
End Sub
Function MYVALID()
MYVALID = True
If xCode.Text = "" Then
    MsgBox "«·ﬂÊœ  ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    MYVALID = False
    Exit Function
End If

If cardtable.RecordCount > 0 And formMode = addmode Then
        cardtable.FindFirst "Code = '" & xCode.Text & "'"
        If Not cardtable.NoMatch Then
            MYVALID = False
            Exit Function
        End If
End If

If xNameA.Text = "" Then
    MYVALID = False
    MsgBox "ÌÃ»  ”ÃÌ· «·√”„"
    Exit Function
End If

If cardtable.RecordCount > 0 And formMode = addmode Then
    cardtable.FindFirst "NAMEA = " & MyParn(xNameA.Text)
    If Not cardtable.NoMatch Then
        MsgBox "Â–« «·»Ì«‰ „ÊÃÊœ „‰ ﬁ»·"
        MYVALID = False
        Exit Function
    End If
End If
End Function
Sub UnEmptyTable()
   CmdFirst.Enabled = True
   CmdLast.Enabled = True
   CmdNext.Enabled = True
   CmdPrevious.Enabled = True
   CmdDel.Enabled = True
   Cmdgo.Enabled = True
   CmdAdd.Enabled = True
   xDescAfilter.Enabled = True
End Sub
Private Sub Cmd_Print_Click()
tempdb.Execute "Delete * From Temp"
Set TempTable = tempdb.OpenRecordset("TEMP")
With cardtable
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
        TempTable.AddNew
        TempTable.str10 = Mid(.NAMEA, 1, 1)
        TempTable.str1 = .NAMEA
        TempTable.str2 = .PHONE1
        TempTable.str3 = .PHONE2
        TempTable.str4 = .PHONE3
        TempTable.STR5 = .PHONE4
        TempTable.str6 = .PHONE5
        TempTable.str7 = .PHONE6
        TempTable.STR19 = firsttitle
        TempTable.str20 = Secondtitle
        TempTable.Update
       .MoveNext
    Loop
End If
End With
myws.BeginTrans
myws.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\phone1.rpt"
Report1.DataFiles(0) = cPathTemp
Report1.Action = 1
End Sub
Private Sub CmdAdd_Click()
AddProc
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    cardtable.Delete
    Set cardtable = mydb.CreateDynaset("phone")
    If cardtable.RecordCount > 0 Then
        cardtable.MoveLast
        MyLoad
    Else
       EmptyForm
    End If
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
cardtable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
    Dim Generalarray(3)
    Dim GrdArray(3)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,NAMEA As «·≈”„, PHONE1 As [ ·Ì›Ê‰] From PHONE"
    Generalarray(3) = "Where NAMEA Like '*cFilter*'"
        
    GrdArray(1) = 1000
    GrdArray(2) = 3500
    GrdArray(3) = 2000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Sub Cmd_OK_Click()
End Sub

Private Sub CmdLast_Click()
cardtable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
cardtable.MoveNext
If cardtable.EOF Then
    cardtable.MovePrevious
Else
    MyLoad
End If
End Sub

Private Sub CmdPhone_Click()
    Dim Generalarray(3)
    Dim GrdArray(8)
        
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Code As «·ﬂÊœ,NAMEA As «·≈”„, PHONE1 , PHONE2, PHONE3, PHONE4, PHONE5, PHONE6 From PHONE"
    Generalarray(3) = "Where PHONE1 Like '*cFilter*' OR PHONE2 Like '*cFilter*' OR PHONE3 Like '*cFilter*' OR PHONE4 Like '*cFilter*' OR PHONE5 Like '*cFilter*' OR PHONE6 Like '*cFilter*' "
        
    GrdArray(1) = 1000
    GrdArray(2) = 2000
    GrdArray(3) = 1000
    GrdArray(4) = 1000
    GrdArray(5) = 1000
    GrdArray(6) = 1000
    GrdArray(7) = 1000
    GrdArray(8) = 1000
        
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Private Sub CmdPrevious_Click()
cardtable.MovePrevious
If cardtable.BOF Then
    cardtable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
If Not MYVALID Then Exit Sub

If Not MsgBox(msgBoxStr, 4) = 6 Then
    CmdUndo_Click
    Exit Sub
End If

Select Case formMode
Case addmode
    cardtable.AddNew
    MyReplace
    cardtable.Update
    cardtable.MoveLast
    AddProc

Case Editmode
    cardtable.Edit
    MyReplace
    cardtable.Update

Case EmptyMode
    cardtable.AddNew
    MyReplace
    cardtable.Update
    cardtable.MoveLast
    AddProc
End Select
End Sub
Private Sub CmdUndo_Click()
Select Case formMode
Case EmptyMode
    myDefine
Case addmode
    cardtable.MoveLast
    editProc
    MyLoad
Case Editmode
    MyLoad
End Select
End Sub
Sub EmptyForm()
    myDefine
    MyMove = False
    formMode = EmptyMode
    xCode.Text = "000001"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If cardtable.RecordCount = 0 Then Exit Sub
If formMode <> Editmode Then Exit Sub
If KeyCode = 34 Then CmdPrevious_Click
If KeyCode = 33 Then CmdNext_Click
End Sub
Private Sub Form_Load()
formMode = Editmode
'Me.Picture = LoadPicture(App.Path & "\DATA.jpg")
'Set myDb = OpenDatabase(App.Path & "\PHONE.mdb")
Set cardtable = mydb.CreateDynaset("select * from PHONE order by code")
'Set TempTable = mydb.CreateDynaset("TEMP")
If cardtable.RecordCount > 0 Then
    MyLoad
    editProc
Else
    EmptyProc
End If
End Sub
Sub Handlecontrols()
Select Case formMode
Case Editmode
     CmdAdd.Enabled = True
     CmdDel.Enabled = True
     CmdInform.Enabled = True
     CmdExit.Enabled = True
     CmdSave.Enabled = True
     CmdUndo.Enabled = True
     CmdPrevious.Enabled = True
     CmdNext.Enabled = True
     CmdLast.Enabled = True
     CmdFirst.Enabled = True
     xCode.Enabled = False
Case addmode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    CmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
Case EmptyMode
    CmdInform.Enabled = False
    CmdDel.Enabled = False
    CmdAdd.Enabled = False
    CmdSave.Enabled = True
    CmdUndo.Enabled = True
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    CmdFirst.Enabled = False
    xCode.Enabled = True
End Select
End Sub
Sub EmptyProc()
formMode = EmptyMode
Handlecontrols
myDefine
xCode.Text = "000001"
End Sub
Sub editProc()
formMode = Editmode
Handlecontrols
End Sub
Sub AddProc()
MyMove = False
formMode = addmode
Handlecontrols
myDefine
xCode.Text = IncRec(myLastField(cardtable, "Code"))
End Sub

