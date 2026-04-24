VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Vs_Cash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   11850
      TabIndex        =   16
      Top             =   0
      Width           =   11910
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Traditional Arabic"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9900
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   75
         Width           =   1815
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   540
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_cash.frx":0000
         Caption         =   "Œ—ÊÃ"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
      End
      Begin Threed.SSCommand CMDDELINV 
         Height          =   540
         Left            =   4912
         TabIndex        =   18
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_cash.frx":0902
         Caption         =   "Õ–› «·„” ‰œ"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmdNewinv 
         Height          =   540
         Left            =   7350
         TabIndex        =   19
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_cash.frx":1154
         Caption         =   "„” ‰œ ÃœÌœ"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   540
         Left            =   2456
         TabIndex        =   20
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_cash.frx":1A92
         Caption         =   "Õ›Ÿ «·„” ‰œ"
         Alignment       =   8
         ButtonStyle     =   2
         PictureAlignment=   6
      End
   End
   Begin VB.TextBox xCode 
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
      Height          =   360
      Left            =   4125
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1125
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton cmd_move 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ﬂ‘› Õ”«»"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4275
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4590
   End
   Begin VB.CommandButton CmdPrevious 
      BackColor       =   &H00E0E0E0&
      Caption         =   "”«»ﬁ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1130
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton CmdNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "·«Õﬁ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2155
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton CmdFirst 
      BackColor       =   &H00E0E0E0&
      Caption         =   "√Ê·"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton CmdLast 
      BackColor       =   &H00E0E0E0&
      Caption         =   "√ŒÌ—"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox xDate 
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
      Height          =   360
      Left            =   4095
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   750
      Width           =   1290
   End
   Begin VB.TextBox xDoc_No 
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
      Height          =   360
      Left            =   8025
      Locked          =   -1  'True
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   2100
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   5310
      Left            =   75
      TabIndex        =   2
      Top             =   2325
      Width           =   11730
      _cx             =   20690
      _cy             =   9366
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   2
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Vs_cash.frx":1EE4
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   1
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin Threed.SSCommand XBAL 
      Height          =   465
      Left            =   9075
      TabIndex        =   15
      Top             =   1800
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   820
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   14654788
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   2
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   675
      Width           =   11715
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·√Ã„«·Ï"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7800
      Width           =   600
   End
   Begin VB.Label LblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   7920
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   2850
      Shape           =   4  'Rounded Rectangle
      Top             =   7875
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ "
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10455
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1155
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·”‰œ"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10455
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   795
      Width           =   930
   End
   Begin VB.Label xClientBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   360
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1155
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ "
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   750
      Width           =   570
   End
End
Attribute VB_Name = "Vs_Cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DocTable As Recordset, ClientTable As Recordset
Dim BalClientTable As Recordset
Dim BoxTable As Recordset
Dim COUNTINVTOTAL As Double
Dim formMode, cSTRMOSM As String
Dim myFileName As String
Dim cFile_10 As String
Dim ClientMoveType  As String
Dim lBox As Boolean
Dim nCrow As Double
Const NewInvMode = 4, applyMode = 5
Sub DocValid()
If xDoc_No.Text = "" Then Exit Sub
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.FindFirst " doc_no = " & MyParn(xDoc_No)
If DocTable.NoMatch Then
    Exit Sub
Else
    ApplyProc
End If
End Sub
Sub EmptyProc()
formMode = EmptyMode
ItemInv.Rows = 1
ItemInv.Rows = 2
End Sub
Sub AddProc()
formMode = addmode
ItemInv.AddItem ""
ItemInv.TextMatrix(ItemInv.Rows - 1, 0) = True
End Sub
Sub Fillgrd()
COUNTINVTOTAL = 0
ItemInv.Rows = 1
i = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow
LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(i, 1) = TurnValue(DocTable!BOX, Null, "")
    
    .TextMatrix(i, 2) = DocTable!CODE
    ClientTable.FindFirst " CODE = " & MyParn(.TextMatrix(i, 2))
    If Not DocTable.NoMatch Then .TextMatrix(i, 3) = ClientTable.DESCA
    .TextMatrix(i, 4) = TurnValue(DocTable.doc, Null, "")
    .TextMatrix(i, 5) = TurnValue(DocTable.DESCA, Null, "")
    .TextMatrix(i, 6) = TurnValue(Format(DocTable!Value, "##0.00"), Null, "")
    If publicFlag = 1 Then .TextMatrix(i, 7) = TurnValue(DocTable!VISA, Null, False)
     COUNTINVTOTAL = COUNTINVTOTAL + DocTable.Value
    LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
    DocTable.MoveNext
    If DocTable.EOF Then Exit Sub
    If DocTable.doc_no <> xDoc_No.Text Then Exit Sub
    i = i + 1
Loop
nCrow = 0
End With
End Sub
Sub CodeLookup()
    ActiveControl.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(2)
    Set Generalarray(1) = Me
    Generalarray(2) = "SELECT CODE AS [—ﬁ„ ], DESCA AS [«·≈”„] FROM " & cFile_10 & " WHERE code is not null "
    Generalarray(3) = " AND DescA Like('*cFilter*')"
    Generalarray(4) = "Order by CODE"
    GrdArray(1) = 1000
    GrdArray(2) = 5000
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
End Sub
Sub ApplyProc()
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If Not DocTable.EOF Then
If DocTable.NoMatch Then
    EmptyProc
    xDoc_No.Enabled = True
Else
    xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
If ActiveControl.Name = ItemInv.Name Then
    ItemInv.EditText = GrdText(Search.Grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 2) = GrdText(Search.Grid1, 0)
    ItemInv.TextMatrix(ItemInv.Row, 3) = GrdText(Search.Grid1, 1)
Else
    ActiveControl.Text = GrdText(Search.Grid1, 0)
End If
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·„” ‰œ"
    MYVALID = False
End If
If xDate.Text = "" Or Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
End Function
Sub Undoinv()
Select Case formMode
Case addmode
    InvGrid.Rows = InvGrid.Rows - 1
    dispProc
Case Editmode
    dispProc
Case EmptyMode
    
End Select
End Sub
Private Sub Cmd_Inv_Click()
xDoc_No.Enabled = False
ItemInv.Enabled = True
ItemInv.SetFocus
ItemInv.Rows = 1
ItemInv.Rows = 2
ItemInv.TextMatrix(1, 0) = True
End Sub
Private Sub cmd_move_Click()
If xCode.Text <> "" Then
    Load ClientMove
    ClientMove.xclient.Text = xCode.Text
    ClientMove.Show 1
End If
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ  »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        xDate.Text = ""
        xClientBalance.Caption = ""
        Fillgrd
        xDoc_No.Enabled = True
        ItemInv.Enabled = False
    End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdNewInv_Click()
If Not MyReplace Then Exit Sub
ItemInv.Rows = 1
ItemInv.AddItem ""
ItemInv.TextMatrix(1, 0) = True

xDate.Text = Date
xClientBalance.Caption = ""
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
Fillgrd
End Sub
Private Sub CmdUndo_Click()
    DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
    If Not DocTable.NoMatch Then
        xDate.Text = Format(DocTable.[Date], "dd-mm-yyyy")
        Fillgrd
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Select Case publicFlag
    Case 1 '  ”œÌœ«  ⁄„·«¡
        myFileName = "File8_10"
        ClientMoveType = "7"
        Set ClientTable = mydb.OpenRecordset("File3_10", dbOpenDynaset)
        Vs_Cash.Caption = " ”œÌœ«  ‰ﬁœÌ… „‰ «·⁄„·«¡"
        cFile_10 = "FILE3_10"
        Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE8_10 ORDER BY DOC_NO ,CODE ", dbOpenDynaset)
    Case 3 '  ”œÌœ«  ⁄„·«¡
        myFileName = "File8_30"
        ClientMoveType = "7"
        Set ClientTable = mydb.OpenRecordset("File4_10", dbOpenDynaset)
        Vs_Cash.Caption = "”œ«œ „Ê—œÌ‰"
        cFile_10 = "FILE4_10"
        Set DocTable = mydb.OpenRecordset("SELECT * FROM FILE8_30 ORDER BY DOC_NO ,CODE ", dbOpenDynaset)
End Select
Set BoxTable = mydb.OpenRecordset("SELECT * FROM file0_50 ORDER BY CODE ", dbOpenDynaset)
lBox = True
If BoxTable.RecordCount = 0 Then lBox = False
cStrBox = StrBox
xDate.Text = Format(Date, "dd-mm-yyyy")
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.doc_no)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .WordWrap = True
    .Cols = 8
    .Rows = 2
    .TextMatrix(1, 0) = True
    .Editable = flexEDKbdMouse
      
    .TextMatrix(0, 1) = "«·Œ“‰…"
    .TextMatrix(0, 2) = "ﬂÊœ"
    .TextMatrix(0, 3) = "«·≈”„"
    .TextMatrix(0, 4) = "„” ‰œ"
    .TextMatrix(0, 5) = "«·»Ì«‰"
    .TextMatrix(0, 6) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 7) = "›Ì“«"
    .ColWidth(0) = 0
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 2500
    .ColWidth(4) = 1000
    .ColWidth(5) = 3500
    .ColWidth(6) = 1300
    .ColWidth(7) = 600
    .ColDataType(0) = 11
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTBoolean
    If publicFlag = 3 Then .ColHidden(7) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColComboList(1) = cStrBox
    .ColHidden(1) = Not lBox
End With
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub ItemInv_EnterCell()
With ItemInv
    If .Row > 0 Then
        xCode.Text = ItemInv.TextMatrix(ItemInv.Row, 2)
        Me.cmd_move.Caption = " ﬂ‘› Õ”«» " & ItemInv.TextMatrix(ItemInv.Row, 3)
    End If
    
    If .Row = .Rows - 1 And .Col = .Cols - 2 Then
        .Rows = .Rows + 1
        .TextMatrix(ItemInv.Rows - 1, 0) = True
    End If
    
    If .Col = 3 Then
        .Editable = flexEDNone
    Else
        .Editable = flexEDKbdMouse
    End If
'    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = &HFFC0C0
    nCrow = .Row
End With
End Sub
Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("Õ–› ”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        ItemInv.RemoveItem ItemInv.Row
    End If
End If
End Sub
Private Sub ItemInv_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case ItemInv.Col
    Case 2
        If KeyCode = 27 Then
            Undoinv
            Exit Sub
        End If
        If KeyCode = 112 Then
            CodeLookup
        End If
End Select
End Sub
Private Sub ItemInv_LeaveCell()
    With ItemInv
        ClientTable.FindFirst " CODE = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 2))
        If Not ClientTable.NoMatch Then
            ItemInv.TextMatrix(ItemInv.Row, 3) = ClientTable.DESCA
            cStr1 = " SELECT Sum(FILE3_11.PAY) AS TPAY, Sum(FILE3_11.SAL) AS TSAL FROM FILE3_11 WHERE CODE = " & MyParn(ItemInv.TextMatrix(ItemInv.Row, 2))
            Set BalClientTable = mydb.OpenRecordset(cStr1)
            
            With BalClientTable
                XBAL.Caption = " —’Ìœ «·Õ”«» " & Format(TurnValue(BalClientTable.TSAL, Null, 0) - TurnValue(BalClientTable.tpay, Null, 0), "#0.00")
            End With
        End If
    End With
End Sub
Private Sub ItemInv_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = False
    If ItemInv.Col <> 2 Then Exit Sub
    If publicFlag = "1" Or publicFlag = "2" Then
        ClientTable.FindFirst " CODE = " & MyParn(ItemInv.EditText)
        If ClientTable.NoMatch Or ItemInv.EditText = "" Then
            Cancel = True
        Else
        End If
    Else
        ClientTable.FindFirst " CODE = " & MyParn(ItemInv.EditText)
        If ClientTable.NoMatch Or ItemInv.EditText = "" Then
            Cancel = True
        End If
    End If
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select " & myFileName & ".Doc_No as [«·„”·”·],[Date] as [ «—ÌŒ ], " & cFile_10 & ".DescA as [«”„ «·⁄„Ì·] " & _
                      " From " & myFileName & " Inner join " & cFile_10 & " on " & myFileName & ".Code = " & cFile_10 & ".code "
    Generalarray(3) = " Where " & cFile_10 & ".DescA Like '*cFilter*' or doc_no Like '*cFilter*'  "
    Generalarray(4) = " ORDER BY " & cFile_10 & ".DescA,DATE "
    GrdArray(1) = 2000
    GrdArray(2) = 1500
    GrdArray(3) = 3000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_LostFocus()
DocValid
End Sub
Function myDelete()
    ' Õ–›  «·„” ‰œ
    cString = " DELETE  " & myFileName & " .* FROM " & myFileName & " WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
        
    ' Õ–› Õ—ﬂ… «·„” ‰œ
    cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(ClientMoveType)
    mydb.Execute cString
    Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ", dbOpenDynaset)
End Function
Function MyReplace()
MyReplace = True
' „—«Ã⁄… «·›« Ê—…
With ItemInv
For i = 1 To .Rows - 1
    ClientTable.FindFirst " CODE = " & MyParn(.TextMatrix(i, 2))
    If ClientTable.NoMatch And .TextMatrix(i, 2) <> "" Then
        .Select i, 0, i, 5
        cMess = "·‰ Ì „  ”ÃÌ· »Ì«‰ «·Œ«’  " & .TextMatrix(i, 2) & " «·ﬂÊœ €Ì— „”Ã· "
        MsgBox cMess
        MyReplace = False
    End If
Next i
End With

If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  " & myFileName & ".*  FROM " & myFileName & "  WHERE " & myFileName & ".DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    '  ”ÃÌ· «·„” ‰œ
    With ItemInv
    
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 2) <> "" Then
        DocTable.AddNew
        DocTable.doc_no = xDoc_No.Text
        DocTable.CASH = True
        DocTable.[Date] = xDate.Text
        DocTable.CODE = .TextMatrix(i, 2)
        DocTable!Value = Val(.TextMatrix(i, 6))
        DocTable.DESCA = TurnValue(.TextMatrix(i, 5), "", Null)
        DocTable.doc = TurnValue(.TextMatrix(i, 4), "", Null)
        If publicFlag = 1 Then DocTable.VISA = TurnValue(.TextMatrix(i, 7), "", False)
        DocTable.BOX = TurnValue(.TextMatrix(i, 1), "", Null)
        DocTable.Update
        End If
    Next i
    End With
    
    ' Õ–› Õ—ﬂ… ⁄„Ì· «·›« Ê—…
    If publicFlag = 1 Or publicFlag = 2 Then
        cString = " DELETE  FILE3_11.* FROM FILE3_11 WHERE FILE3_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(ClientMoveType)
        mydb.Execute cString
    Else
        cString = " DELETE  FILE4_11.* FROM FILE4_11 WHERE FILE4_11.DOC_ID = " & MyParn(xDoc_No.Text) & " AND [TYPE] = " & MyParn(ClientMoveType)
        mydb.Execute cString
    End If
    
    ' ≈‰‘«¡ Õ—ﬂ… ⁄„Ì· ··›« Ê—…
    
    Select Case publicFlag

        Case 1
            cString = "Insert Into File3_11(" & _
                      "[Type],Doc_Id,Code,[Date],Pay,DescA)" & _
                      " Select '7',Doc_No,Code,[Date],[Value],'œ›⁄…' " & _
                      " From File8_10" & _
                        " where doc_no = " & MyParn(xDoc_No.Text)
        Case 2
            cString = "Insert Into File3_11(" & _
                      "[Type],Doc_Id,Code,[Date],Pay,DescA)" & _
                      " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ… ' " & _
                      " From File8_10" & _
                        " where doc_no = " & MyParn(xDoc_No.Text)
        Case 3
            cString = "Insert Into File4_11(" & _
                      "[Type],Doc_Id,Code,[Date],Pay,DescA)" & _
                      " Select '7',Doc_No,Code,[Date],[Value],'œ›⁄…' " & _
                      " From File8_30" & _
                        " where doc_no = " & MyParn(xDoc_No.Text)
        Case 4
            cString = "Insert Into File4_11(" & _
                      "[Type],Doc_Id,Code,[Date],Pay,DescA)" & _
                      " Select '0',Doc_No,Code,[Date],[Value],' ”ÊÌ… ' " & _
                      " From File8_40" & _
                        " where doc_no = " & MyParn(xDoc_No.Text)
    
    End Select
    mydb.Execute cString
    Set DocTable = mydb.OpenRecordset("SELECT * FROM " & myFileName & " ORDER BY DOC_NO ", dbOpenDynaset)
End If
End Function
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 6 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub CmdFirst_Click()
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.MoveFirst
xDoc_No.Text = DocTable.doc_no
DocValid
End Sub
Private Sub CmdLast_Click()
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.MoveLast
xDoc_No.Text = DocTable.doc_no
DocValid
End Sub
Private Sub CmdNext_Click()
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.FindLast " DOC_NO = " & MyParn(xDoc_No)
DocTable.MoveNext
If Not DocTable.EOF Then
    xDoc_No.Text = DocTable.doc_no
    DocValid
End If
End Sub
Private Sub CmdPrevious_Click()
If DocTable.RecordCount = 0 Then Exit Sub
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No)
DocTable.MovePrevious
If Not DocTable.BOF Then
    xDoc_No.Text = DocTable.doc_no
    DocValid
End If
End Sub
Private Function StrBox()
If BoxTable.RecordCount > 0 Then
    BoxTable.MoveFirst
    i = 1
    StrBox = "#  " & ";       "
    StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!DESCA
    BoxTable.MoveNext
    Do While True
        i = i + 1
        If BoxTable.EOF Then Exit Do
        StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!DESCA
        BoxTable.MoveNext
    Loop
End If
End Function

