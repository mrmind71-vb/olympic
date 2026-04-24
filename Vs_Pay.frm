VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Vs_Pay 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„”ÕÊ»«  ‘—ﬂ«¡"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   9900
      TabIndex        =   13
      Top             =   0
      Width           =   9960
      Begin Threed.SSCommand CmdExit 
         Height          =   540
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         BackColor       =   14737632
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_Pay.frx":0000
         Caption         =   "Œ—ÊÃ"
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Threed.SSCommand CMDDELINV 
         Height          =   540
         Left            =   4912
         TabIndex        =   15
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         BackColor       =   14737632
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_Pay.frx":0902
         Caption         =   "Õ–› «·„” ‰œ"
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Threed.SSCommand cmdNewinv 
         Height          =   540
         Left            =   7368
         TabIndex        =   16
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         BackColor       =   14737632
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_Pay.frx":1154
         Caption         =   "„” ‰œ ÃœÌœ"
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   540
         Left            =   2456
         TabIndex        =   17
         Top             =   0
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   953
         _Version        =   196610
         Font3D          =   2
         ForeColor       =   192
         BackColor       =   14737632
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Vs_Pay.frx":1A92
         Caption         =   "Õ›Ÿ «·„” ‰œ"
         Alignment       =   4
         PictureAlignment=   1
      End
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
      Height          =   390
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1965
      Width           =   765
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
      Height          =   390
      Left            =   2055
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1965
      Width           =   765
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
      Height          =   390
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1965
      Width           =   765
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
      Height          =   390
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1965
      Width           =   765
   End
   Begin VB.CommandButton CmdUndo 
      BackColor       =   &H00E0E0E0&
      Caption         =   " —«Ã⁄"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   375
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1425
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   4005
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   765
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
      Left            =   9150
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1590
   End
   Begin VSFlex7LCtl.VSFlexGrid ItemInv 
      Height          =   5085
      Left            =   75
      TabIndex        =   2
      Top             =   2445
      Width           =   11595
      _cx             =   20452
      _cy             =   8969
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
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
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Vs_Pay.frx":1EE4
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
   Begin Threed.SSCommand xClosed 
      Height          =   390
      Left            =   6030
      TabIndex        =   12
      Top             =   720
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonStyle     =   4
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   1260
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   645
      Width           =   11655
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈Ã„«·Ï"
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
      Height          =   195
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   7755
      Width           =   705
   End
   Begin VB.Label LblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   7755
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7635
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "—ﬁ„ „” ‰œ"
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
      Left            =   10875
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   735
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
      Left            =   5370
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   750
      Width           =   570
   End
End
Attribute VB_Name = "Vs_Pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cStrBox As String
Dim DocTable As Recordset
Dim BoxTable As Recordset
Dim COUNTINVTOTAL As Double
Dim ManTable As Recordset
Dim formMode, cStrMan As String
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
End Sub
Sub Fillgrd()
COUNTINVTOTAL = 0
ItemInv.Rows = 1
I = 1
With ItemInv
.FixedRows = 1
.ExplorerBar = flexExSortShow
LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then Exit Sub
Do While True
   .AddItem ""
    .TextMatrix(I, 0) = TurnValue(DocTable!MAN, Null, "")
    .TextMatrix(I, 1) = TurnValue(DocTable!DESCA, Null, "")
    .TextMatrix(I, 2) = TurnValue(Format(DocTable!Value, "###0.00"), Null, "")
    .TextMatrix(I, 3) = TurnValue(DocTable!BOX, Null, "")
     COUNTINVTOTAL = COUNTINVTOTAL + DocTable!Value
    LblTotal.Caption = Format(COUNTINVTOTAL, "##0.00")
    DocTable.MoveNext
    If DocTable.EOF Then Exit Sub
    If UCase(DocTable!DOC_NO) <> UCase(xDoc_No.Text) Then Exit Sub
    I = I + 1
Loop
End With
End Sub
Sub ApplyProc()
If Not DocTable.EOF Then
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No.Text)
If DocTable.NoMatch Then
    EmptyProc
    xDoc_No.Enabled = True

Else
    xDate.Text = Format(DocTable![Date], "dd-mm-yyyy")
    Fillgrd
    dispProc
    xDoc_No.Enabled = False
End If
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Function MYVALID()
MYVALID = True
If xDoc_No.Text = "" Then
    MsgBox " ”ÃÌ· —ﬁ„ «·„” ‰œ"
    MYVALID = False
End If
If Not IsDate(xDate.Text) Then
    MsgBox " ”ÃÌ· «· «—ÌŒ"
    MYVALID = False
End If
End Function
Sub Undoinv()
Select Case formMode
    Case addmode
        invGrid.Rows = invGrid.Rows - 1
        dispProc
    Case Editmode
        dispProc
    Case EmptyMode
End Select
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ  »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        myDelete
        xDoc_No.Text = ""
        xDate.Text = ""
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
ItemInv.Rows = 2
ItemInv.AddItem ""
xDate.Text = Date
xDoc_No.Enabled = True
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable.DOC_NO)
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
        xDate.Text = Format(DocTable![Date], "dd-mm-yyyy")
        Fillgrd
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Set DocTable = mydb.OpenRecordset("SELECT * FROM File8_90  order by DATE , doc_no  ", dbOpenDynaset)
Set ManTable = mydb.OpenRecordset("SELECT * FROM file1_70 WHERE FLAG = 5 ORDER BY CODE ", dbOpenDynaset)
Set BoxTable = mydb.OpenRecordset("SELECT * FROM file0_50  ORDER BY CODE ", dbOpenDynaset)
cStrMan = StrMan
cStrBox = StrBox
xDate.Text = Format(Date, "dd-mm-yyyy")
If DocTable.RecordCount > 0 Then
    DocTable.MoveLast
    xDoc_No.Text = IncRec(DocTable!DOC_NO)
Else
    xDoc_No.Text = "000001"
End If
With ItemInv
    .Cols = 4
    .Rows = 2
    .Editable = flexEDKbdMouse
    
    .TextMatrix(0, 0) = "«·‘—Ìﬂ"
    .TextMatrix(0, 1) = "«·»Ì«‰"
    .TextMatrix(0, 2) = "«·≈Ã„«·Ï"
    .TextMatrix(0, 3) = "«·Œ“‰…"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 1000
    .ColWidth(3) = 2000
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTDouble
    
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColComboList(0) = cStrMan
    .ColComboList(3) = cStrBox
End With
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub ItemInv_EnterCell()
With ItemInv
    If .Row = .Rows - 1 Then .Rows = .Rows + 1
End With
End Sub
Private Sub ItemInv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("Õ–› ”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        ItemInv.RemoveItem ItemInv.Row
    End If
End If
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xDoc_No.Text = ""
    Dim Generalarray(4)
    Dim GrdArray(3)
    Set Generalarray(1) = Me
    Generalarray(2) = "Select Doc_No as [«·„”·”·],format([Date],'dd-mm-yyyy') as [ «—ÌŒ ], DescA as [»Ì«‰] " & _
                      " From  File8_90 "
    Generalarray(3) = " Where DescA Like '%cFilter%' or doc_no Like '%cFilter%' "
    Generalarray(4) = " ORDER BY DATE "
    GrdArray(1) = 1000
    GrdArray(2) = 1500
    GrdArray(3) = 4000
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Show 1
End If
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_No.Text = UCase(xDoc_No.Text)
DocValid
End Sub
Function myDelete()
    ' Õ–›  «·„” ‰œ
    cString = " DELETE  File8_90 FROM File8_90  WHERE File8_90.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    DocTable.Requery
End Function
Function MyReplace()
MyReplace = True
' „—«Ã⁄… «·„” ‰œ
With ItemInv

If MyReplace Then
    ' Õ–› «·„” ‰œ ﬁ»· «· ⁄œÌ·
    cString = " DELETE  File8_90   FROM File8_90  WHERE  File8_90.DOC_NO = " & MyParn(xDoc_No.Text)
    mydb.Execute cString
    
    '  ”ÃÌ· «·„” ‰œ
    
    For I = 1 To .Rows - 1
        If .TextMatrix(I, 0) <> "" Then
        DocTable.AddNew
        DocTable!DOC_NO = xDoc_No.Text
        DocTable![Date] = xDate.Text
        DocTable!Value = Val(.TextMatrix(I, 2))
        DocTable!MAN = TurnValue(.TextMatrix(I, 0), "", Null)
        DocTable!BOX = TurnValue(.TextMatrix(I, 3), "", Null)
        DocTable!DESCA = TurnValue(.TextMatrix(I, 1), "", Null)
        DocTable.Update
        End If
    Next I
    DocTable.Requery
End If
    End With
End Function
Private Function StrMan()
If ManTable.RecordCount > 0 Then
    ManTable.MoveFirst
    I = 1
    StrMan = "#  " & ";       "
    StrMan = StrMan & "|#" & ManTable!CODE & ";" & ManTable!DESCA
    ManTable.MoveNext
    Do While True
        I = I + 1
        If ManTable.EOF Then Exit Do
        StrMan = StrMan & "|#" & ManTable!CODE & ";" & ManTable!DESCA
        ManTable.MoveNext
    Loop
End If
End Function
Private Sub ItemInv_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If ItemInv.Col = 4 Then
    KeyAscii = RetNumber(KeyAscii, True)
End If
End Sub
Private Sub CmdFirst_Click()
DocTable.MoveFirst
xDoc_No.Text = DocTable!DOC_NO
DocValid
End Sub
Private Sub CmdLast_Click()
DocTable.MoveLast
xDoc_No.Text = DocTable!DOC_NO
DocValid
End Sub
Private Sub CmdNext_Click()
DocTable.FindLast " DOC_NO = " & MyParn(xDoc_No)
DocTable.MoveNext
If Not DocTable.EOF Then
    xDoc_No.Text = DocTable!DOC_NO
    DocValid
End If
End Sub
Private Sub CmdPrevious_Click()
DocTable.FindFirst " DOC_NO = " & MyParn(xDoc_No)
DocTable.MovePrevious
If Not DocTable.BOF Then
    xDoc_No.Text = DocTable!DOC_NO
    DocValid
End If
End Sub
Private Function StrBox()
If BoxTable.RecordCount > 0 Then
    BoxTable.MoveFirst
    I = 1
    StrBox = "#  " & ";       "
    StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!DESCA
    BoxTable.MoveNext
    Do While True
        I = I + 1
        If BoxTable.EOF Then Exit Do
        StrBox = StrBox & "|#" & BoxTable!CODE & ";" & BoxTable!DESCA
        BoxTable.MoveNext
    Loop
End If
End Function

